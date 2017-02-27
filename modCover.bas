Attribute VB_Name = "modCover"
Option Explicit

Public Const DVD_BotY As Integer = 199 'нижняя координата двд DVD_Height-15
Public Const DVD_Height As Integer = 184 'нижняя координата двд

'for rot font
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const LOGPIXELSY As Integer = 90        '  Logical pixels/inch in Y


Public Sub DrawRotatedText(ByVal txt As String, _
                           ByVal X As Single, ByVal Y As Single, _
                           ByVal font_name As String, ByVal Size As Long, _
                           ByVal weight As Long, ByVal escapement As Long, _
                           ByVal use_italic As Boolean, ByVal use_underline As Boolean, _
                           ByVal use_strikethrough As Boolean, ByVal fColor As Long, _
                           ByVal fcharset As Long, ByVal wID As Long, Optional centr As Boolean = False)

''used with fdwCharSet
'Const ANSI_CHARSET = 0
'Const DEFAULT_CHARSET = 1
'Const SYMBOL_CHARSET = 2
'Const SHIFTJIS_CHARSET = 128
'Const HANGEUL_CHARSET = 129
'Const CHINESEBIG5_CHARSET = 136
'Const OEM_CHARSET = 255
''used with fdwOutputPrecision
'Const OUT_CHARACTER_PRECIS = 2
Const OUT_DEFAULT_PRECIS = 0
'Const OUT_DEVICE_PRECIS = 5
''used with fdwClipPrecision
'Const CLIP_DEFAULT_PRECIS = 0
'Const CLIP_CHARACTER_PRECIS = 1
'Const CLIP_STROKE_PRECIS = 2
''used with fdwQuality
'Const DEFAULT_QUALITY = 0
'Const DRAFT_QUALITY = 1
Const PROOF_QUALITY = 2
''used with fdwPitchAndFamily
Const DEFAULT_PITCH = 0
'Const FIXED_PITCH = 1
'Const VARIABLE_PITCH = 2

Const CLIP_LH_ANGLES = 16   ' Needed for tilted fonts.
'Const PI = 3.14159625
'Const PI_180 = PI / 180#

Dim newFont As Long
Dim OldFont As Long
Dim temp As Long
Dim trt As Boolean
Dim hwnd As Long
Dim hdc As Long
hwnd = GetDesktopWindow
hdc = GetDC(hwnd)

'    newfont = CreateFont(size, 0, _
     escapement, escapement, weight, _
     use_italic, use_underline, _
     use_strikethrough, fcharset, OUT_DEFAULT_PRECIS, _
     CLIP_LH_ANGLES, PROOF_QUALITY, DEFAULT_PITCH, font_name)

newFont = CreateFont(-(Size * GetDeviceCaps(hdc, LOGPIXELSY)) / 72, 0, _
                     escapement, escapement, weight, _
                     use_italic, use_underline, _
                     use_strikethrough, fcharset, OUT_DEFAULT_PRECIS, _
                     CLIP_LH_ANGLES, PROOF_QUALITY, DEFAULT_PITCH, font_name)

With FrmMain
    ' Select the new font.
    OldFont = SelectObject(.PicCoverPaper.hdc, newFont)


    ' Display the text.
    '.PicCoverPaper.CurrentX = x
    '.PicCoverPaper.CurrentY = y

    temp = .PicCoverPaper.ForeColor
    .PicCoverPaper.ForeColor = fColor    '&HFF00&
    'Debug.Print PicCoverPaper.TextWidth(txt)

    Do While .PicCoverPaper.TextWidth(txt) >= wID    'длина текста
        trt = True
        txt = left$(txt, Len(txt) - 6)
    Loop

    If trt Then
        txt = txt + "...": trt = False
    End If
    'Else
    Dim yy As Long
    
        If centr Then
        If escapement = 900 Then
            yy = Y - (wID - .PicCoverPaper.TextWidth(txt)) / 2 '- 2
            If yy < Y Then Y = yy
        Else
        'сверху вниз
            yy = Y + (wID - .PicCoverPaper.TextWidth(txt)) / 2 '+ 2
            If .TabStripCover.SelectedItem.Index = 3 Or .TabStripCover.SelectedItem.Index = 4 Then yy = yy + 2 'коррекция для двд (большой шрифт)
            If yy > Y Then Y = yy
        End If
        End If
    'End If
    
    .PicCoverPaper.CurrentX = X
    .PicCoverPaper.CurrentY = Y
    
    .PicCoverPaper.Print txt
    .PicCoverPaper.ForeColor = temp

    ' Restore the original font.
    newFont = SelectObject(.PicCoverPaper.hdc, OldFont)
    ' Free font resources (important!)
    DeleteObject newFont
    
End With
End Sub

Public Sub ShowCoverConvert()
Dim temp As String, temp2 As String
Dim annot As String
Dim sFnt As StdFont
'Dim tempWes As Integer
Dim PRatio As Currency
Dim tmpb As Boolean

On Error Resume Next
If rs.RecordCount < 1 Then Exit Sub

With FrmMain
    .PicCoverPaper.ScaleMode = 6    'мм
    .PicCoverPaper.Width = 11940
    .PicCoverPaper.Height = 16860
    .PicCoverPaper.ScaleWidth = 210.6085
    .PicCoverPaper.ScaleHeight = 297.392

    '                                       Название или Метка
    If (.chPrnAllOne.Value = vbChecked) And (CheckCount > 1) Then
        If rs.Fields("Label").Value <> vbNullString Then
            temp = rs.Fields("Label").Value
        Else
            temp = vbNullString
        End If
    Else
        If rs.Fields("MovieName").Value <> vbNullString Then
            temp = rs.Fields("MovieName").Value
        Else
            temp = vbNullString
        End If
    End If

    If Not IsNull(rs.Fields("Annotation").Value) Then annot = rs.Fields("Annotation").Value

    Set .PicCoverPaper = Nothing
    .PicCoverPaper.CurrentX = 1: .PicCoverPaper.CurrentY = 1
    Set .PicCoverPaper.Font = FontHor
    .PicCoverPaper.Print "Sur Video Catalog: " & temp

    Set sFnt = FontVert

    'закраска фона аннотации
    .PicCoverPaper.Line (35, 25)-(155, 145), CoverHorBackColor, BF

    'название и аннотация
    Set .PicCoverTextWnd = Nothing
    'потом PicCoverTextWnd.Visible = True
    .PicCoverTextWnd.BackColor = CoverHorBackColor
    temp2 = .PicCoverTextWnd.ForeColor
    .PicCoverTextWnd.ForeColor = HFontColor    '&HFF00&

    Dim txtBody As String
    If .chPrnAllOne.Value = vbChecked Then
        txtBody = GetCoverSpisok
    Else
        If Len(temp) = 0 Then
            txtBody = annot
        Else
            txtBody = temp & " " & vbCrLf & vbCrLf & " " & annot
        End If
    End If

    WrapText .PicCoverTextWnd, txtBody, 1, .PicCoverTextWnd.Width, 1, .PicCoverTextWnd.Height, False   'без рамки

    '''''''
    tmpb = GetPic(.PicCoverTemp, 1, "SnapShot1")
    If tmpb Then
        .PicCoverTemp.Picture = .PicCoverTemp.Image
        PRatio = .PicCoverTemp.Width / .PicCoverTemp.Height    '.PicTempHid(0).Height
        If PRatio < 1 Then PRatio = 1.333
        .PicCoverPaper.PaintPicture .PicCoverTemp.Picture, 35.4, 25.5, 39.5, 39.5 / PRatio
    End If

    Set .PicCoverTemp = Nothing
    If tmpb Then
        tmpb = GetPic(.PicCoverTemp, 1, "SnapShot2")
        .PicCoverTemp.Picture = .PicCoverTemp.Image
        PRatio = .PicCoverTemp.Width / .PicCoverTemp.Height    '.PicTempHid(0).Height
        If PRatio < 1 Then PRatio = 1.333
        .PicCoverPaper.PaintPicture .PicCoverTemp.Picture, 75.3, 25.5, 39.5, 39.5 / PRatio
    End If

    Set .PicCoverTemp = Nothing
    tmpb = GetPic(.PicCoverTemp, 1, "SnapShot3")
    If tmpb Then
        .PicCoverTemp.Picture = .PicCoverTemp.Image
        PRatio = .PicCoverTemp.Width / .PicCoverTemp.Height    '.PicTempHid(0).Height
        If PRatio < 1 Then PRatio = 1.333
        .PicCoverPaper.PaintPicture .PicCoverTemp.Picture, 115.4, 25.5, 39.5, 39.5 / PRatio
    End If

    Set .PicCoverTemp = Nothing
    ''''''''''''

    .PicCoverTextWnd.ForeColor = temp2
    .PicCoverTextWnd.Picture = .PicCoverTextWnd.Image
    .PicCoverPaper.PaintPicture .PicCoverTextWnd.Picture, .PicCoverTextWnd.left, .PicCoverTextWnd.top


    'обложка
    'If .PicFaceV.Picture = 0 Then 'всегда надо, если несколько печатаем то не обновится обложка, или проверять флаг
    GetPic .PicFaceV, 1, "FrontFace"
    If (Not NoPicFrontFaceFlag) And (.ChPrintPix.Value = 1) Then PutPrintPixConvert

    .PicCoverPaper.DrawWidth = 1

    Call PrintBoxesConvert

    .PicCoverPaper.Picture = .PicCoverPaper.Image


    .FrameCover.Visible = True

    'показать окно текста
    .PicCoverTextWnd.Visible = True

    .PicPrintScroll_Resize

End With
End Sub

Public Sub ShowCoverStandard()
Dim temp As String, temp2 As String
Dim annot As String
Dim sFnt As StdFont
Dim tempWes As Integer
'Dim temp As String
Dim PRatio As Currency
Dim tmpb As Boolean
Dim tmpC As Currency

On Error Resume Next
With FrmMain

    .PicCoverPaper.ScaleMode = 6    'мм
    'PicCoverPaper.Move 0, 0, 11940, 16860 - делает ресайз
    .PicCoverPaper.Width = 11940
    .PicCoverPaper.Height = 16860
    .PicCoverPaper.ScaleWidth = 210.6085
    .PicCoverPaper.ScaleHeight = 297.392

    '                                       Название или Метка
    If .chPrnAllOne.Value = vbChecked And CheckCount > 1 Then
        If rs.Fields("Label").Value <> vbNullString Then
            temp = rs.Fields("Label").Value
        Else
            temp = vbNullString
        End If
    Else
        If rs.Fields("MovieName").Value <> vbNullString Then
            temp = rs.Fields("MovieName").Value
        Else
            temp = vbNullString
        End If
    End If

    If Not IsNull(rs.Fields("Annotation").Value) Then annot = rs.Fields("Annotation").Value

    Set .PicCoverPaper = Nothing
    .PicCoverPaper.CurrentX = 1: .PicCoverPaper.CurrentY = 1
    Set .PicCoverPaper.Font = FontHor
    Set .PicCoverTextWnd.Font = FontHor

    .PicCoverPaper.Print "Sur Video Catalog: " & temp

    'закраска фона аннотации
    .PicCoverPaper.Line (29, 145)-(178, 262), CoverVertBackColor, BF
    .PicCoverPaper.Line (35, 145)-(172, 262), CoverHorBackColor, BF


    Set sFnt = FontVert
    If sFnt.Bold Then tempWes = FW_BOLD Else tempWes = FW_NORMAL

    DrawRotatedText temp, 29, 260, _
                    sFnt.name, sFnt.Size, _
                    tempWes, 90 * 10, _
                    sFnt.Italic, sFnt.Underline, sFnt.Strikethrough, VFontColor, sFnt.Charset, 113, Check2Bool(.ChCentrTitle.Value)


    DrawRotatedText temp, 178, 147, _
                    sFnt.name, sFnt.Size, _
                    tempWes, 270 * 10, _
                    sFnt.Italic, sFnt.Underline, sFnt.Strikethrough, VFontColor, sFnt.Charset, 113, Check2Bool(.ChCentrTitle.Value)

    'название и аннотация
    temp2 = .PicCoverTextWnd.ForeColor
    .PicCoverTextWnd.ForeColor = HFontColor    '&HFF00&
    '******************************************
    'PicCoverTextWnd.BorderStyle = 0
    Set .PicCoverTextWnd = Nothing
    'потом PicCoverTextWnd.Visible = True

    .PicCoverTextWnd.BackColor = CoverHorBackColor
    .PicCoverTextWnd.Line (0, 0)-(.PicCoverTextWnd.Width, .PicCoverTextWnd.Height), CoverHorBackColor, BF

    Dim txtBody As String
    If .chPrnAllOne.Value = vbChecked Then
        txtBody = GetCoverSpisok
    Else
        txtBody = temp & " " & vbCrLf & vbCrLf & " " & annot
    End If

    WrapText .PicCoverTextWnd, txtBody, 1, .PicCoverTextWnd.Width, 0, .PicCoverTextWnd.Height - 4, False

'''''''''''''''''''
    tmpb = GetPic(.PicCoverTemp, 1, "SnapShot1")
    If tmpb Then
        .PicCoverTemp.Picture = .PicCoverTemp.Image
        PRatio = .PicCoverTemp.Width / .PicCoverTemp.Height
        If PRatio < 1 Then PRatio = 1.333
        .PicCoverPaper.PaintPicture .PicCoverTemp.Picture, 35.4, 145.5, 45.2, 45.2 / PRatio
    End If

    Set .PicCoverTemp = Nothing
    tmpb = GetPic(.PicCoverTemp, 1, "SnapShot2")
    If tmpb Then
        .PicCoverTemp.Picture = .PicCoverTemp.Image
        PRatio = .PicCoverTemp.Width / .PicCoverTemp.Height
        If PRatio < 1 Then PRatio = 1.333
        .PicCoverPaper.PaintPicture .PicCoverTemp.Picture, 80.9, 145.5, 45.2, 45.2 / PRatio
    End If

    Set .PicCoverTemp = Nothing
    tmpb = GetPic(.PicCoverTemp, 1, "SnapShot3")
    If tmpb Then
        .PicCoverTemp.Picture = .PicCoverTemp.Image
        PRatio = .PicCoverTemp.Width / .PicCoverTemp.Height
        If PRatio < 1 Then PRatio = 1.333
        tmpC = 45.2 / PRatio
        .PicCoverPaper.PaintPicture .PicCoverTemp.Picture, 126.5, 145.5, 45.2, tmpC
    End If
    
    Set .PicCoverTemp = Nothing
        
    '*************************************
    
    .PicCoverTextWnd.ForeColor = temp2
    .PicCoverTextWnd.Picture = .PicCoverTextWnd.Image
    .PicCoverPaper.PaintPicture .PicCoverTextWnd.Picture, .PicCoverTextWnd.left, .PicCoverTextWnd.top

    'обложка
    'If .PicFaceV.Picture = 0 Then 'всегда надо, если несколько печатаем то не обновится обложка, или проверять флаг
    GetPic .PicFaceV, 1, "FrontFace"
    If NoPicFrontFaceFlag Or .ChPrintPix.Value = vbUnchecked Then
    Else
        PutPrintPixStandard
    End If

    .PicCoverPaper.DrawWidth = 1

    'каркас
    Call PrintBoxesStandard
    .PicCoverPaper.Picture = .PicCoverPaper.Image

    .FrameCover.Visible = True

    'показать окно текста
    .PicCoverTextWnd.Visible = True

    .PicPrintScroll_Resize

End With
End Sub

Public Sub ShowCoverSpisok()
Dim temp2 As String    ', temp2 As String

With FrmMain
    .PicCoverPaper.ScaleMode = 6    'мм
    'PicCoverPaper.Move 0, 0, 11940, 16860 - делает ресайз
    .PicCoverPaper.Width = 11940
    .PicCoverPaper.Height = 16860
    .PicCoverPaper.ScaleWidth = 210.6085
    .PicCoverPaper.ScaleHeight = 297.392


    Set .PicCoverPaper = Nothing
    .PicCoverPaper.CurrentX = 1: .PicCoverPaper.CurrentY = 1
    Set .PicCoverPaper.Font = FontHor
    Set .PicCoverTextWnd.Font = FontHor

    .PicCoverPaper.Print "Sur Video Catalog"    ': " & temp

    'закраска фона аннотации
    'PicCoverPaper.Line (29, 145)-(178, 262), CoverVertBackColor, BF
    'PicCoverPaper.Line (35, 145)-(172, 262), CoverHorBackColor, BF

    'название и аннотация
    temp2 = .PicCoverTextWnd.ForeColor
    .PicCoverTextWnd.ForeColor = HFontColor    '&HFF00&

    '******************************************

    .PicCoverTextWnd.BorderStyle = 1

    Set .PicCoverTextWnd = Nothing
    'PicCoverTextWnd.Visible = True 'потом

    .PicCoverTextWnd.BackColor = CoverHorBackColor
    .PicCoverTextWnd.Line (0, 0)-(.PicCoverTextWnd.Width, .PicCoverTextWnd.Height), CoverHorBackColor, BF

    WrapText .PicCoverTextWnd, GetCoverSpisok, 0, .PicCoverTextWnd.Width, 0, .PicCoverTextWnd.Height - 4, False
    'WrapText PicCoverTextWnd, GetCoverSpisok, 10, PicCoverTextWnd.Width, 10, PicCoverTextWnd.Height - 14, False

    .PicCoverTextWnd.ForeColor = temp2

    .PicCoverTextWnd.Picture = .PicCoverTextWnd.Image
    .PicCoverPaper.PaintPicture .PicCoverTextWnd.Picture, .PicCoverTextWnd.left, .PicCoverTextWnd.top

    Set .PicCoverTemp = Nothing

    .PicCoverPaper.DrawWidth = 1

    'каркас
    'Call PrintBoxesStandard
    .PicCoverPaper.Picture = .PicCoverPaper.Image

    .FrameCover.Visible = True
    'показать окно текста
    .PicCoverTextWnd.Visible = True

    .PicPrintScroll_Resize
End With
End Sub

Public Sub ShowCoverDVD(slim As Boolean)
Dim temp As String, temp2 As String
Dim annot As String
Dim sFnt As StdFont
Dim tempWes As Integer
Dim tmpb As Boolean
Dim PRatio As Currency

On Error Resume Next
With FrmMain

    .PicCoverPaper.ScaleMode = 6    'мм
    .PicCoverPaper.Width = 16860
    .PicCoverPaper.Height = 11940
    .PicCoverPaper.ScaleWidth = 297.392
    .PicCoverPaper.ScaleHeight = 210.6085

    '                                       Название или Метка
    If .chPrnAllOne.Value = vbChecked And CheckCount > 1 Then
        If rs.Fields("Label").Value <> vbNullString Then
            temp = rs.Fields("Label").Value
        Else
            temp = vbNullString
        End If
    Else
        If rs.Fields("MovieName").Value <> vbNullString Then
            temp = rs.Fields("MovieName").Value
        Else
            temp = vbNullString
        End If
    End If

    If Not IsNull(rs.Fields("Annotation").Value) Then annot = rs.Fields("Annotation").Value

    Set .PicCoverPaper = Nothing
    .PicCoverPaper.CurrentX = 1: .PicCoverPaper.CurrentY = 1
    Set .PicCoverPaper.Font = FontHor
    Set .PicCoverTextWnd.Font = FontHor

    .PicCoverPaper.Print "Sur Video Catalog: " & temp

    'закраска фона аннотации

    Set sFnt = FontVert
    If sFnt.Bold Then tempWes = FW_BOLD Else tempWes = FW_NORMAL



    If slim Then
        .PicCoverPaper.Line (10, 15)-(140, DVD_BotY), CoverHorBackColor, BF
        .PicCoverPaper.Line (140, 15)-(148, DVD_BotY), CoverVertBackColor, BF
        .PicCoverPaper.Line (148, 15)-(278, DVD_BotY), CoverHorBackColor, BF

        DrawRotatedText temp, 148, 18, _
                        sFnt.name, sFnt.Size + 3, _
                        tempWes, 270 * 10, _
                        sFnt.Italic, sFnt.Underline, sFnt.Strikethrough, VFontColor, sFnt.Charset, 170, Check2Bool(.ChCentrTitle.Value)
    Else
        .PicCoverPaper.Line (10, 15)-(140, DVD_BotY), CoverHorBackColor, BF
        .PicCoverPaper.Line (140, 15)-(153, DVD_BotY), CoverVertBackColor, BF
        .PicCoverPaper.Line (153, 15)-(283, DVD_BotY), CoverHorBackColor, BF

        DrawRotatedText temp, 151, 18, _
                        sFnt.name, sFnt.Size + 6, _
                        tempWes, 270 * 10, _
                        sFnt.Italic, sFnt.Underline, sFnt.Strikethrough, VFontColor, sFnt.Charset, 170, Check2Bool(.ChCentrTitle.Value)
    End If



    'название и аннотация
    temp2 = .PicCoverTextWnd.ForeColor
    .PicCoverTextWnd.ForeColor = HFontColor    '&HFF00&

    '******************************************

    Set .PicCoverTextWnd = Nothing
    ' потом PicCoverTextWnd.Visible = True

    .PicCoverTextWnd.BackColor = CoverHorBackColor
    .PicCoverTextWnd.Line (0, 0)-(.PicCoverTextWnd.Width, .PicCoverTextWnd.Height), CoverHorBackColor, BF

    Dim txtBody As String
    If .chPrnAllOne.Value = vbChecked Then
        txtBody = GetCoverSpisok
    Else
        txtBody = temp & " " & vbCrLf & vbCrLf & " " & annot
    End If

    WrapText .PicCoverTextWnd, txtBody, 1, .PicCoverTextWnd.Width, 0, .PicCoverTextWnd.Height - 4, False

'''''''''''
    tmpb = GetPic(.PicCoverTemp, 1, "SnapShot1")
    If tmpb Then
        .PicCoverTemp.Picture = .PicCoverTemp.Image
        PRatio = .PicCoverTemp.Width / .PicCoverTemp.Height    '.PicTempHid(0).Height
        If PRatio < 1 Then PRatio = 1.333
        .PicCoverPaper.PaintPicture .PicCoverTemp.Picture, 12, 17, 63, 63 / PRatio
    End If

    Set .PicCoverTemp = Nothing
    tmpb = GetPic(.PicCoverTemp, 1, "SnapShot2")
    If tmpb Then
        .PicCoverTemp.Picture = .PicCoverTemp.Image
        PRatio = .PicCoverTemp.Width / .PicCoverTemp.Height    '.PicTempHid(0).Height
        If PRatio < 1 Then PRatio = 1.333
        .PicCoverPaper.PaintPicture .PicCoverTemp.Picture, 76, 17, 63, 63 / PRatio
    End If

    Set .PicCoverTemp = Nothing

    '*************************************
    .PicCoverTextWnd.ForeColor = temp2
    .PicCoverTextWnd.Picture = .PicCoverTextWnd.Image
    .PicCoverPaper.PaintPicture .PicCoverTextWnd.Picture, .PicCoverTextWnd.left, .PicCoverTextWnd.top


    'обложка
    'If .PicFaceV.Picture = 0 Then 'всегда надо, если несколько печатаем то не обновится обложка, или проверять флаг
    GetPic .PicFaceV, 1, "FrontFace"
    If NoPicFrontFaceFlag Or .ChPrintPix.Value = vbUnchecked Then
    Else
        If slim Then
            PutPrintPixDVD slim:=True
        Else
            PutPrintPixDVD slim:=False
        End If
    End If

    .PicCoverPaper.DrawWidth = 1

    'каркас
    If slim Then
        Call PrintBoxesDVDSlim
    Else
        Call PrintBoxesDVD
    End If

    .PicCoverPaper.Picture = .PicCoverPaper.Image
    .FrameCover.Visible = True

    'показать окно текста
    .PicCoverTextWnd.Visible = True

    .PicPrintScroll_Resize
End With
End Sub

Public Sub PutPrintPixConvert()
Dim PRatio As Double
Dim WIDTH1 As Long, HEIGHT1 As Long

'Dim ph As Long, pw As Long
On Error Resume Next 'вдруг ошибка в виа
If FrmMain.ChPrintPix.Value = vbUnchecked Then Exit Sub

With FrmMain

    .PicCoverPaper.PaintPicture .ImBlankHid.Image, 35, 145, 120, 120
    .PicFaceV.Picture = .PicFaceV.Image

    If .PicFaceV.Picture <> 0 Then

        PRatio = .PicFaceV.Height / .PicFaceV.Width
        If PRatio > 1 Then PRatio = 1 / PRatio

        If .ChPropP.Value = vbChecked Then    'сохранять пропорции
            If .ChScaleP.Value = vbChecked Then    'растянуть
                If .PicFaceV.Height > .PicFaceV.Width Then
                    If .ChCentrP.Value = vbChecked Then
                        'centre hor
                        .PicCoverPaper.PaintPicture .PicFaceV.Picture, 155 - 60 + (120 * PRatio) / 2, 265, -120 * PRatio, -120
                    Else
                        .PicCoverPaper.PaintPicture .PicFaceV.Picture, 155, 265, -120 * PRatio, -120
                    End If
                Else
                    If .ChCentrP.Value = vbChecked Then
                        'centre vert
                        .PicCoverPaper.PaintPicture .PicFaceV.Picture, 155, 265 - 60 + (120 * PRatio) / 2, -120, -120 * PRatio
                    Else
                        .PicCoverPaper.PaintPicture .PicFaceV.Picture, 155, 265, -120, -120 * PRatio
                    End If
                End If
            Else    'не растягивать 'сохранять проп
                If .ChCentrP.Value = vbChecked Then
                    .PicCoverPaper.PaintPicture .PicFaceV.Picture, 155, 265, -120, -120, (.PicFaceV.ScaleWidth - 120) / 2, (.PicFaceV.ScaleHeight - 120) / 2, 120, 120
                Else
                    .PicCoverPaper.PaintPicture .PicFaceV.Picture, 155, 145 + .PicFaceV.ScaleHeight, -.PicFaceV.ScaleWidth, -.PicFaceV.ScaleHeight
                End If
            End If    'ChScaleP
        Else    'не сохранять пропорции
            If .ChScaleP.Value = vbChecked Then    'растянуть
                '.PicCoverPaper.PaintPicture .PicFaceV.Picture, 155, 265, -120, -120    'перевернуть
                WIDTH1 = .ScaleX(120, vbMillimeters, vbPixels)
                ResizeWIA .PicCoverTemp, WIDTH1, WIDTH1, .PicFaceV, rot:=180
                .PicCoverPaper.PaintPicture .PicCoverTemp.Picture, 35, 145

            End If    'ChScaleP
        End If    'ChPropP

        'закрыть вылезающие за края полная большая картинка
        .ImBlankHid.BackColor = .PicCoverPaper.BackColor
        .PicCoverPaper.PaintPicture .ImBlankHid.Image, 35, 265, 125, 125    'вниз
        .PicCoverPaper.PaintPicture .ImBlankHid.Image, 0, 145, 35, 200    'влево
        .ImBlankHid.BackColor = CoverHorBackColor
    End If

    Call PrintBoxesConvert
End With
End Sub

Public Sub PutPrintPixStandard()
Dim PRatio As Double
Dim WIDTH1 As Long, HEIGHT1 As Long

On Error Resume Next 'вдруг ошибка в виа
If FrmMain.ChPrintPix.Value = vbUnchecked Then Exit Sub

With FrmMain
    .PicCoverPaper.PaintPicture .ImBlankHid.Image, 35, 25, 120, 120             '< размер
    .PicFaceV.Picture = .PicFaceV.Image

    If .PicFaceV.Picture <> 0 Then

        PRatio = .PicFaceV.Height / .PicFaceV.Width
        If PRatio > 1 Then PRatio = 1 / PRatio

        If .ChPropP.Value = vbChecked Then  'сохранять пропорции
            If .ChScaleP.Value = vbChecked Then  'растянуть
                If .PicFaceV.Height > .PicFaceV.Width Then
                    If .ChCentrP.Value = vbChecked Then
                        'centre по гориз
                        .PicCoverPaper.PaintPicture .PicFaceV.Picture, 35 + 60 - (120 * PRatio) / 2, 25, 120 * PRatio, 120
                    Else
                        .PicCoverPaper.PaintPicture .PicFaceV.Picture, 35, 25, 120 * PRatio, 120
                    End If

                Else
                    If .ChCentrP.Value = vbChecked Then
                        'centre по верт
                        .PicCoverPaper.PaintPicture .PicFaceV.Picture, 35, 25 + 60 - (120 * PRatio) / 2, 120, 120 * PRatio
                    Else
                        .PicCoverPaper.PaintPicture .PicFaceV.Picture, 35, 25, 120, 120 * PRatio
                    End If

                End If
            Else    'не растягивать
                If .ChCentrP.Value = vbChecked Then
                    ' Centre полюбому (кроп)
                    'PicFaceV дб в милиметрах
                    .PicCoverPaper.PaintPicture .PicFaceV.Picture, 35, 25, , , (.PicFaceV.ScaleWidth - 120) / 2, (.PicFaceV.ScaleHeight - 120) / 2, 120, 120
                Else
                    .PicCoverPaper.PaintPicture .PicFaceV.Picture, 35, 25, , , , , 120, 120
                End If
            End If    'ChScaleP
            
        Else    'не сохранять пропорции
            If .ChScaleP.Value = vbChecked Then  'растянуть
                '.PicCoverPaper.PaintPicture .PicFaceV.Picture, 35, 25, 120, 120
                WIDTH1 = .ScaleX(120, vbMillimeters, vbPixels)
                HEIGHT1 = WIDTH1 '.ScaleY(25, vbMillimeters, vbPixels)
                ResizeWIA .PicCoverTemp, WIDTH1, HEIGHT1, .PicFaceV
                .PicCoverPaper.PaintPicture .PicCoverTemp.Picture, 35, 25

            End If    'ChScaleP
        End If    'ChPropP


    End If

    Call PrintBoxesStandard
End With
End Sub

Public Sub PutPrintPixDVD(slim As Boolean)
Dim PRatio As Double
Dim k As Double
Dim X1 As Single, Y1 As Single
Dim WIDTH1 As Long, HEIGHT1 As Long
Dim chOr As Boolean

Dim PicX1 As Long

If FrmMain.ChPrintPix.Value = vbUnchecked Then Exit Sub
On Error Resume Next 'вдруг ошибка в виа

If slim Then
    PicX1 = 148
Else
    PicX1 = 153
End If

With FrmMain
    .PicCoverPaper.PaintPicture .ImBlankHid.Image, PicX1 + 0.2, 15.3, 130, DVD_Height - 0.3
    .PicFaceV.Picture = .PicFaceV.Image



    If .PicFaceV.Picture <> 0 Then
        PRatio = .PicFaceV.Height / .PicFaceV.Width
        k = DVD_Height / 130
        If k < PRatio Then chOr = True
        If chOr Then If PRatio > 1 Then PRatio = 1 / PRatio

        'PicCoverTemp
        '< размер
        If .ChPropP.Value = vbChecked Then  'сохранять пропорции
            If .ChScaleP.Value = vbChecked Then  'растянуть
                If chOr Then
                    X1 = (128 * PRatio * k)
                    WIDTH1 = .ScaleX(X1, vbMillimeters, vbPixels)
                    HEIGHT1 = .ScaleY(DVD_Height, vbMillimeters, vbPixels)
                    ResizeWIA .PicCoverTemp, WIDTH1, HEIGHT1, .PicFaceV
                    If .ChCentrP.Value = vbChecked Then
                        'centre hor
                        .PicCoverPaper.PaintPicture .PicCoverTemp.Picture, PicX1 + 65 - X1 / 2, 15
                    Else
                        .PicCoverPaper.PaintPicture .PicCoverTemp.Picture, PicX1, 15
                    End If

                Else
                    Y1 = DVD_Height * PRatio / k
                    WIDTH1 = .ScaleX(130, vbMillimeters, vbPixels)
                    HEIGHT1 = .ScaleY(Y1, vbMillimeters, vbPixels)
                    ResizeWIA .PicCoverTemp, WIDTH1, HEIGHT1, .PicFaceV
                    If .ChCentrP.Value = vbChecked Then
                        'centre VERT
                        .PicCoverPaper.PaintPicture .PicCoverTemp.Picture, PicX1 + 0.2, 15 + 90 - Y1 / 2
                    Else
                        .PicCoverPaper.PaintPicture .PicCoverTemp.Picture, PicX1 + 0.2, 15
                    End If
                End If

            Else    'не растягивать
                If .ChCentrP.Value = vbChecked Then
                    .PicCoverPaper.PaintPicture .PicFaceV.Picture, PicX1, 15, , , (.PicFaceV.ScaleWidth - 130) / 2, (.PicFaceV.ScaleHeight - DVD_Height) / 2, 130, DVD_Height
                Else
                    .PicCoverPaper.PaintPicture .PicFaceV.Picture, PicX1, 15, , , , , 130, DVD_Height
                End If
            End If    'ChScaleP

        Else    'не сохранять пропорции
            WIDTH1 = .ScaleX(130, vbMillimeters, vbPixels)
            HEIGHT1 = .ScaleY(DVD_Height, vbMillimeters, vbPixels)
            ResizeWIA .PicCoverTemp, WIDTH1, HEIGHT1, .PicFaceV
            If .ChScaleP.Value = vbChecked Then  'растянуть
                .PicCoverPaper.PaintPicture .PicCoverTemp.Picture, PicX1 + 0.2, 15
            End If    'ChScaleP
        End If    'ChPropP
    End If

    Set .PicCoverTemp = Nothing
    
If slim Then
    PrintBoxesDVDSlim
Else
    PrintBoxesDVD
End If

End With
End Sub

Private Sub PrintBoxesConvert()
With FrmMain
    .PicCoverPaper.Line (35, 25)-(155, 145), , B
    .PicCoverPaper.Line (35, 145)-(155, 265), , B

    'PicCoverPaper.Line (29, 145)-(178, 262), , B
    'PicCoverPaper.Line (35, 145)-(172, 262), , B
End With
End Sub
Private Sub PrintBoxesStandard()
With FrmMain
    .PicCoverPaper.Line (35, 25)-(155, 145), , B

    .PicCoverPaper.Line (29, 145)-(178, 262), , B
    .PicCoverPaper.Line (35, 145)-(172, 262), , B
End With
End Sub

Private Sub PrintBoxesDVD()
With FrmMain
    .PicCoverPaper.Line (10, 15)-(140, DVD_BotY), , B  '1  w=140-10=130                    h=195-15=180
    .PicCoverPaper.Line (140, 15)-(153, DVD_BotY), , B '2               w=153-140=13       h=180
    .PicCoverPaper.Line (153, 15)-(283, DVD_BotY), , B '3 w=283-153=130                    h=180
End With
End Sub

Private Sub PrintBoxesDVDSlim()
With FrmMain
    .PicCoverPaper.Line (10, 15)-(140, DVD_BotY), , B '1  w=140-10=130                    h=195-15=180
    .PicCoverPaper.Line (140, 15)-(148, DVD_BotY), , B 'w=148-140=8
    .PicCoverPaper.Line (148, 15)-(278, DVD_BotY), , B 'w=278-148=130
End With
End Sub
Public Function WrapText(Pict As Object, ByVal txt As String, ByVal xmin As Single, ByVal xmax As Single, ByVal ymin As Single, ByVal ymax As Single, ByVal draw_box As Boolean) As String
' Print the text with wrapping.
'Dim x As Single
'Dim y As Single
Dim xmargin As Single
Dim ymargin As Single
Dim line_wid As Single
Dim new_line As String
Dim new_word As String
Dim FirstLineFlag As Boolean
'Dim CRFlag As Boolean
'Dim oldfntbld As bolean

FirstLineFlag = True
NonPrintToSpace txt


' If we should draw a box, add a small margin.
If draw_box Then
    xmargin = Pict.TextWidth("x") / 2
    ymargin = Pict.ScaleY(Pict.Font.Size * 0.5, vbPoints, Pict.ScaleMode)
    xmin = xmin + xmargin
    xmax = xmax - xmargin
    ymin = ymin + ymargin
End If
'Debug.Print txt
line_wid = xmax - xmin

' Start printing.
Pict.CurrentY = ymin
Pict.CurrentX = xmin
new_word = GetWord(txt)
Do
    ' Start with the last word examined.
    ' Note that this loop prints at least one
    ' word per line. That is important if the
    ' text contains a word too long to fit on
    ' a line.
    new_line = new_word

    Do

        ' Get the next word.
        new_word = GetWord(txt)
        If LenB(new_word) = 0 Then Exit Do
        'If InStr(1, new_word, vbCr) Then Pict.Print

        '            CRFlag = False
        If InStr(1, new_word, vbCr) Then
            ReplaceCharacters new_word, vbCr, vbNullString
            '            CRFlag = True
            If Pict.TextWidth(new_line & " " & new_word) > line_wid Then Exit Do
            new_line = new_line & " " & new_word
            new_word = vbNullString
            Exit Do
        End If
        ' See if the new word fits.

        If Pict.TextWidth(new_line & " " & new_word) > line_wid Then Exit Do

        If (Pict.CurrentY) > ymax Then Exit Do

        'борьба с (enter & фраза) - печаталось с самого левого края
        If left$(new_word, 1) = vbLf Then
            If Len(new_word) > 1 Then
                new_word = vbLf & Space(xmin) & right$(new_word, Len(new_word) - 1)
            End If

        End If
        new_line = new_line & " " & new_word
    Loop
    If (Pict.CurrentY) > ymax Then Exit Do



    ' Display the line. This moves CurrentX to
    ' zero and CurrentY to the next line.

    Pict.Print new_line

    If FirstLineFlag Then WrapText = new_line: FirstLineFlag = False
    If LenB(txt) = 0 And right$(new_line, Len(new_word)) = new_word Then Exit Do

    ' Reset CurrentX to our left margin.
    Pict.CurrentX = xmin

Loop

' Draw the box if desired.
If draw_box Then
    xmin = xmin - xmargin
    xmax = xmax + xmargin
    ymin = ymin - ymargin
    Pict.Line (xmin, ymin)-(xmax, Pict.CurrentY + ymargin), , B
End If
End Function

Private Function GetCoverSpisok() As String
'mnuCopyLV_Click
'mnuExportCheckClip_Click
'GetCoverSpisokLabel
'GetCoverSpisok
'Export2Excel

If FrmMain.chPrnAllOne.Value = vbChecked Then
    ' If (LVSortColl > 0) Or (LVSortColl = -1) Then    'если сортировано
    '    GetCoverSpisok = GetCoverSpisokLabel
    '    Exit Function
    ' End If
Else
    GetCoverSpisok = GetCoverSpisokLabel
    Exit Function
End If

' все помеченные подряд, пронумерованное
Dim j As Integer, M As Integer
Dim temp As String, temp2 As String
'Dim doflag As Boolean
Dim addLabel As Boolean
Dim tmp As String, tmpm As String
Dim pArr() As Integer
Dim allArr() As String

Screen.MousePointer = vbHourglass

With FrmMain
    'заполнить массис соответствия   индекс в списке - позиция
    'индекс массива - позиция, значение - индекс поля списка

    ReDim pArr(.ListView.ColumnHeaders.Count) 'As Integer
    For j = 0 To LstExport_ListCount     'все 0-24, потом индекс не обработаем
        pArr(.ListView.ColumnHeaders(j + 1).Position) = j
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
                    If CheckCount > 1 Then
                        addLabel = True
                    Else
                        addLabel = False
                    End If
                Else
                    addLabel = False
                    tmpm = vbNullString    ' и не надо боле
                End If
            End If
        End If


        'поля
        For j = 1 To .ListView.ColumnHeaders.Count   '1-25 индекс не обработаем

            If LstExport_Arr(pArr(j)) Then

                '                'одинаковы ли метки
                '                If M = 1 And j = 1 Then    'каждая запись каждое начало цикла
                '                    tmpm = CheckNoNullVal(dbLabelInd)
                '                Else
                '                    If tmpm <> vbNullString Then
                '                        If CheckNoNullVal(dbLabelInd) = tmpm Then
                '                            If CheckCount > 1 Then
                '                                addLabel = True
                '                            Else
                '                                addLabel = False
                '                            End If
                '                        Else
                '                            addLabel = False
                '                            tmpm = vbNullString    ' и не надо боле
                '                        End If
                '                    End If
                '                End If

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
                If IsNotEmptyOrZero(tmp) Then temp = temp & " " & tmp
            Else
                'иначе разделитель
                If IsNotEmptyOrZero(tmp) Then temp = temp & ExportDelim & tmp
            End If
        End If

        If Len(temp) <> 0 Then
            If CheckCount > 1 Then
                temp = M & ". " & temp    ' пункт списка
            Else
                temp = LTrim$(temp)
            End If
            'temp2 = temp2 & temp & vbCrLf
            allArr(M) = temp & vbCrLf
        End If
        temp = vbNullString

        'не делать много
        If M > 150 Then Exit For
    Next M
    temp2 = Join(allArr, vbNullString)

    'восстановить текущ. поз. в базе
    RestoreBasePos

    If addLabel Then temp2 = temp2 & vbCrLf & frmEditor.LFilm(1) & ": " & tmpm

    If Len(temp2) = 0 Then temp2 = NamesStore(6)    'Совет: выберите нужные поля для списка в настройках экспорта.
    GetCoverSpisok = temp2

    Screen.MousePointer = vbNormal
End With
End Function
Private Function GetCoverSpisokLabel() As String
'mnuCopyLV_Click
'mnuExportCheckClip_Click
'GetCoverSpisokLabel
'GetCoverSpisok

' список группирован по сортированному полю, если нет сортировки - простой список
'Метка: 10 DVD+R (4)
'1) Библиотекарь. В поисках копья (The Librarian: Quest for the Spear);Приключения;2004
'2) Больше, чем любовь, A Lot Like Love;Комедия, Мелодрама;2005
'3) Влюбись в меня, если осмелишься, Детские игры (Jeux d'enfants);романтическая комедия;2005
'4) Дэнни - цепной пес (Unleashed, Danny the Dog);Боевик, Драма;2005

Dim strSQL As String
Dim rsTmp As DAO.Recordset
Dim grf As String
Dim coun As Integer, i As Integer, j As Integer, k As Integer
'Dim doflag As Boolean
Dim temp As String, temp2 As String
'Dim addLabel As Boolean
Dim tmp As String    ', tmpm As String
Dim NextGrFieldFlag As Boolean
Dim oldGrField As String
Dim GrField As String
Dim AscDesc As String    ' порядок сортировки текущий LVSortOrder
Dim pArr() As Integer
Dim allArr() As String

Screen.MousePointer = vbHourglass

With FrmMain

    If (LVSortColl > 0) And (LVSortColl <> lvHeaderIndexPole) Then    'если сортировано и не индексом

        'заполнить массив соответствия   индекс в списке - позиция
        'индекс массива - позиция, значение - индекс поля списка

        ReDim pArr(.ListView.ColumnHeaders.Count) 'As Integer
        For j = 0 To LstExport_ListCount     'все 0-24, потом индекс не обработаем
            pArr(.ListView.ColumnHeaders(j + 1).Position) = j
        Next j

        ReDim allArr(UBound(CheckRows))

        grf = GetGroupFieldName(LVSortColl)
        
        If Len(grf) <> 0 Then
            'strSQL = "SELECT " & grf & " FROM STORAGE Group by " & grf
            'strSQL = "SELECT * FROM STORAGE WHERE " & grf & " IN (" & strSQL & ") AND Checked = '1' Order By " & grf

            If LVSortOrder = lvwAscending Then
                AscDesc = "Asc"
            Else
                AscDesc = "Desc"
            End If
'сортировать по груп-полю, потом по названию
            strSQL = "SELECT * FROM STORAGE WHERE Checked = '1' Order By " & grf & " " & AscDesc & " , MovieName Asc"
            

           ' On Error Resume Next

            Set rsTmp = DB.OpenRecordset(strSQL)

            If rsTmp.RecordCount > 0 Then

                rsTmp.MoveLast: rsTmp.MoveFirst
                If rsTmp.RecordCount > 150 Then coun = 150 Else coun = rsTmp.RecordCount    'не делать много
                NextGrFieldFlag = True    'первый раз писать поле группировки

                For i = 0 To coun - 1    'по строкам

                       '     If j = 1 Then
                       'пункт списка
                                oldGrField = GrField
                                If Not IsNull(rsTmp(grf)) Then GrField = UCase$(Trim$(rsTmp(grf))) Else GrField = "Empty"
                                If GrField = oldGrField Then
                                    k = k + 1
                                Else
                                    k = 1
                                End If
                        '    End If

                    'поля
                    For j = 1 To .ListView.ColumnHeaders.Count   '1-25 индекс не обработаем


                        If LstExport_Arr(pArr(j)) Then

'                            If j = 1 Then
'                                oldGrField = GrField
'                                If Not IsNull(rsTmp(grf)) Then GrField = UCase$(Trim$(rsTmp(grf))) Else GrField = "Empty"
'
'                                If GrField = oldGrField Then
'                                    k = k + 1    'пункт списка
'                                Else
'                                    k = 1
'                                End If
'                            End If

                            If pArr(j) <> dbAnnotationInd Then
                            
                                If pArr(j) = dbFileNameInd Then
                                    tmp = GetFNamesForSpisok(CheckNoNullValMyRS(pArr(j), rsTmp))
                                Else
                                    tmp = CheckNoNullValMyRS(pArr(j), rsTmp)
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
                        If Not IsNull(rsTmp(dbAnnotationInd)) Then tmp = rsTmp(dbAnnotationInd) Else tmp = vbNullString

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
                            temp = k & ". " & temp    ' пункт списка

                        Else
                            temp = LTrim$(temp)
                        End If
                        'temp2 = temp2 & temp & vbCrLf
                        allArr(i) = temp & vbCrLf
                    End If


                    'строка с полем группировки

                    If GrField <> oldGrField Then
                        allArr(i) = vbCrLf & "  " & GrField & vbCrLf & allArr(i)
                    End If

                    temp = vbNullString

                    rsTmp.MoveNext
                Next i    'строки
                
                temp2 = Join(allArr, vbNullString)

                rsTmp.Close: Set rsTmp = Nothing

            End If            'rsTmp.RecordCount > 0
        End If        'grf
        
    Else    'не сортировано
        'Совет: выберите нужные поля для списка в настройках экспорта.
        'Сортируйте список по нужному для группировки полю
        If Len(temp2) = 0 Then temp2 = NamesStore(9)

    End If    'сортировано

End With
Screen.MousePointer = vbNormal
GetCoverSpisokLabel = temp2
End Function


Private Sub NonPrintToSpace(ByRef txt As String)
' Convert non-printable characters into spaces.
Dim i As Long
Dim txtlen As Long
Dim ch As String

txtlen = Len(txt)
For i = 1 To txtlen
    ch = Mid$(txt, i, 1)
    '      If ch < " " Then Mid$(txt, i, 1) = " "
    'If ch < vbCr Then Mid$(txt, i, 1) = " "
    If ch = vbCr Then Mid$(txt, i, 1) = " "    'vbCrLf
    'Or ch > "~"

Next i
'Debug.Print txt
End Sub

Public Function GetFNamesForSpisok(s As String) As String
Dim a() As String
Dim i As Integer
Dim tmps As String

If Len(s) = 0 Then Exit Function
If Opt_FileWithPath Then
    GetFNamesForSpisok = s 'Replace(s, " |", ",")
Else 'вернуть без пути
    If InStr(s, "|") > 0 Then
        If Tokenize04(s, a(), "|", False) > -1 Then
            For i = 0 To UBound(a)
                tmps = tmps & GetNameFromPathAndName(a(i)) & ", "
            Next i
            GetFNamesForSpisok = left$(tmps, Len(tmps) - 2)
        End If
    Else
        'один
        GetFNamesForSpisok = GetNameFromPathAndName(s)
    End If

End If

End Function
Private Function Check2Bool(ch_val As Integer) As Boolean
'валью чекбокса в булин
If ch_val = vbChecked Then Check2Bool = True
End Function
