Attribute VB_Name = "MainMod"
Option Explicit
Option Compare Text

'Private b1() As Byte 'for simil
'Private b2() As Byte 'for simil

'Public m_cDib As New cDIBSection 'jpeg
Public msgsvc(52) As String 'массив сообщений об ошибках     и вопросов                    ERRORS
'массив хранящий нужные для интерфейса слова, не попавшие в имена контролов
Public NamesStore(12) As String 'Name Stores
Public Const SVCBaseFielsCount = 31 '25 'полей в базе фильмов
Public Const ComboSitesNum = 50 '20 ' скоко пунктов сайтов
Public Const GenreCount = 50 '26 'комбик жанра
Public Const CountryCount = 50 '20 'комбик стран
Public Const ComboNosNum = 50 'комбик носителя
Public Const cbTotal = 10 'кол-во комбиков в фильтре

Public Const Kavs = """"

Public LastMovieFolder As String 'папка последнего открытого файла в ploaddialog
Public UCLVShowPersonFlag As Boolean 'карточка показывает персону, а не обложку

Public arr_chAF(5) As Integer 'галочки фильтра актеров
Public arr_chAD(0) As Integer

Public ActWWWsite As String 'сайт для поиска актера по кнопке (google ...)

Public NoResizePlease As Boolean

Public sPerson As String 'текущая персона выделенная для поиска (TextItemHid_MouseUp UCLV_tActMouseUp)

'Public MovieForCuptureOpened As Boolean 'OpenAddmovFlag флаг , что фильм открыт только для захвата кадра
Public ChangeFromCode_optAspect As Boolean 'флаг клика через код

Public LCID As Long

Public ShowCoverFlag As Boolean 'показано ли окно обложки

Public extAvi As String 'расшерения файлов форматов
Public extDS As String
Public ExtPix As String
Public ExtTxt As String

Public TranslatedFieldsNames() As String 'переведенные названия полей
Public FormShowPicIsModal As Boolean
Public ShowPicFocus As Boolean 'нет при клике на LVClick, да - при клике в 1:1
Public SShotsCount As Integer 'колво доступных скриншотов в текущей записи
    
Public LastBaseIsGood As Boolean 'переходить на предыд базу, если текущая - говно

Public ScrShotEd_W As Long 'размер окна скриншотов и видео в редакторе, начальные
Public ScrShotEd_H As Long
Public MovieEd_W As Long
Public MovieEd_H As Long
Public MovieHeight As Single 'скейлед
Public MovieWidth As Single

'шрифты
Public FontListView As New StdFont
Public FontVert As New StdFont
Public FontHor As New StdFont

Public LastImageListInd As Integer 'сколько иконок в  ImageList, в коде дополняется nopic в конец
                                    'If ImageList.ListImages.Count >= 3 Then
                                    
'для формы поиск/замена
Public Enum LV_AllSelCheck 'где в списке искать
    AllLVRows = 0
    SelectedLVRows = 1
    CheckedLVRows = 2
End Enum
Public Enum AnyWholeFirst
    Search_Anywhere = 0
    Search_WholeField = 1
    Search_StartWith = 2
    Search_EndWith = 3
    Search_Shablon = 4
End Enum
Public Enum BeginEnd
    sBegin = 0
    sEnd = 1
End Enum
Public Enum HowConvert
    LCaseAll = 0
    UCaseAll = 1
    UcaseFirst = 2
    UCaseWord = 3
End Enum


Public LastKey As Long 'ключ (lvclick)

'Public LastTVitem As String 'куда кликали в группировке
Public GroupColumnHeader As String 'название первой колонки списка групп

Public LastSQLGroupString As String 'последний групповой запрос      в скобочках!
Public LastSQLFilterString As String 'последний запрос фильтрации
Public LastSQLPersonString As String 'от фильтров по актеру (из карточки и окна актеров)


Public GroupField As String ' текущее имя поля группировки
Public GroupInd As Integer ' текущий индекс (dbMovieNameInd) поля группировки 'в начале установлен в -1

Public AutoShots As Boolean 'делать проверочный скриншот


Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public AutoNoMessFlag As Boolean 'не задавать вопросов
'Public SameDiskFlag As Boolean 'не спрашивать, если диск уже в базе
Public NoVideoProcess As Boolean 'не возится с рендером (ави)

Public Mark2SaveFlag As Boolean 'да, редактируется поле редактора

Public MeWidth As Long 'из ини
Public MeHeight As Long

Public Const EM_CANUNDO = &HC6
Public Const EM_UNDO = &HC7
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_UNDO = &H304

'Public Const WM_SETTEXT = &HC
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long

'показывать контент при таскании формы
Private Const SPI_GETDRAGFULLWINDOWS = 38
Private Const SPI_SETDRAGFULLWINDOWS = 37
Private Const SPIF_SENDWININICHANGE = &H2
'mzt Private Const SPIF_UPDATEINIFILE = &H1
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
      (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private m_bSwitchOff As Boolean
''''''''''''''''''''''''''''''''''''''

Public SelCount As Long 'кол-во выделенных фильмов
Public CheckCount As Long ' помеченных v

'размеры текстов в Обложке
Public Type Razmer
l As Single
t As Single
w As Single
H As Single
End Type
Public cov_stan As Razmer
Public cov_conv As Razmer
Public cov_dvd As Razmer
Public cov_list As Razmer

Public LVSortOrder As Integer 'текущий сортордер
Public LVSortColl As Integer 'текущая сортированная колонка = -1 если по чекбоксам
'Public LVSortHeader As Integer 'lv сортировано по колонке
'Public LVSortChecked As Boolean 'lv сортировано по помеченному v

'search combo
Public Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Const CB_FINDSTRINGEXACT = &H158
Private Const CB_FINDSTRING = &H14C
'search listbox
Private Const LB_FINDSTRINGEXACT = &H1A2
Private Const LB_FINDSTRING = &H18F

'text
Public Const WM_SETREDRAW = &HB
'
Public pwd As String 'пароль на текущ базу
Public LVActerFilled As Boolean ' таблица актеров заполнена
Public oldTabLVInd As Integer 'текущий таб LV

'чистка памяти
'Public Declare Sub CoFreeUnusedLibraries Lib "ole32" ()
'Private Declare Function GetCurrentProcess Lib "KERNEL32" () As Long
'Private Declare Function SetProcessWorkingSetSize Lib "KERNEL32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long


'
'for Background
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type
Public lngBrush As Long
Public rctMain As RECT
''''''''

''''
'Public LstFilesShowFlag As Boolean 'флаг показан или нет LstFiles - выбор файла фильма для просмотра

Public LVWidth As Single '% ширины LV
Public TVWidth As Single '% ширины TV
Public SplitLVD As Single 'ширина левого окна нижнего сплиттера
Public SSCoverAnnotW As Single
Public SSCoverAnnotH As Single
Public SSCoverAnnotT As Single

Public PicManualFlag As Boolean 'как появилось frmShowPic кликом или ...

Public MaxScroll As Integer 'для прокрутки в FormShowPic

Public exitNukeflag As Boolean 'флаг ошибки в Нюке
Public lvItemLoaded() As Boolean 'c 1? массив флагов загрузки LV, false - только названия загружены у текущей строки

Public frmMainFlag As Boolean 'загружена ли форма
Public frmPeopleFlag As Boolean
Public frmFilterFlag As Boolean
Public frmBinFlag As Boolean
Public SplashFlag As Boolean 'загружен ли Сплэшскрин
'Public FormShowPicFlag As Boolean 'показан ли ПикСкрин
Public FormShowPicLoaded As Boolean 'загружен ли ПикСкрин
Public frmOptFlag As Boolean
Public frmAutoFlag As Boolean
Public frmSRFlag As Boolean
Public frmActFiltFlag  As Boolean
Public frmEditorFlag As Boolean 'виден редактор или нет


Public ViewScrShotFlag As Boolean 'что показывает FormShowPic
Public FilterActFlag As Boolean 'фильтрован ли список актеров

'таймер
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public TicScrSaver As Long 'тайм дла скринсейвера

Public lastHTMLfolderPath As String 'путь для экспорта
Public lastAutoAddfolderPath As String 'путь для автозаполнения

Public userFile As String 'файл user.lng с полным путем

'Public SShotClickFlag As Boolean 'для слайдера, что кликнули из Image0

Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function GlobalAlloc Lib "kernel32.dll" (Optional ByVal Flags As Long = &H42&, Optional ByVal Length As Long = 0&) As OLE_HANDLE
Public Declare Function GlobalFree Lib "kernel32.dll" (ByVal hGlobal As OLE_HANDLE) As Long
Public Declare Function GlobalLock Lib "kernel32.dll" (ByVal hGlobal As OLE_HANDLE) As Long
Public Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hGlobal As OLE_HANDLE) As Long
Public Declare Function OpenClipboard Lib "user32.dll" (Optional ByVal hwnd As OLE_HANDLE = 0&) As Long
Public Declare Function CloseClipboard Lib "user32.dll" () As Long
Public Declare Function EmptyClipboard Lib "user32.dll" () As Long
Public Declare Function SetClipboardData Lib "user32.dll" (Optional ByVal Format As VBRUN.ClipBoardConstants = vbCFDIB, Optional ByVal hGlobal As OLE_HANDLE = 0&) As OLE_HANDLE

'Public LastBaseInd As Integer 'индекс текущей базы в списке (для запоминания при выходе) = CurrentBaseIndex
Public CoverMoveDirection As Integer ' куда растягиваем текстовое окно в обложке Список
Public ChPrintCheckedFlag As Integer 'запомнить значение флажка
'Public ChPrintAllOneFlag As Integer 'там в клике процесс

'Public LVPopupFlag As Boolean 'Видно глюконуло ... Не позиционировать LV по правому клику

Public DebugFileFlagRW As Boolean ' открыт ли файл дебага на запись
Public GlobalFileFlagRW As Boolean ' true - можно писать файл global.ini
Public INIFileFlagRW As Boolean ' ' true - можно писать файл *.ini

Public getPeopleFlag As Boolean 'различать OpenURLPic для людей и редактора

Public DB As DAO.Database
Public rs As DAO.Recordset
'Public rsf As DAO.Recordset
Public ADB As DAO.Database
Public ars As DAO.Recordset
Public DBdd As DAO.Database 'для драг дропа
Public rsdd As DAO.Recordset
Public rsJoin As DAO.Recordset 'для объединения в списке


Public MouseOverTabLV As Integer 'над какой кнопкой мышь в табстрипе выбора баз в окне LV

Public SVCDflag As Boolean 'флаг MPV2 svcd
Public MPGCodec As String 'MPV1 , MPV2 или что вернет mediainfo.dll
'Public LastIndIni As Integer 'курсор на последнее кликнутое поле из ини файла

Public DrDroFlag As Boolean 'обновлять ковер после драг дропа

Public FilteredFlag As Boolean 'флаг неполного показа в lv
Public GroupedFlag As Boolean 'флаг неполного показа в lv
Public FiltPersonFlag As Boolean '---
Public FiltValidationFlag As Boolean '--- метка.серийник

Public FrameViewCaption As String
Public FrameActerCaption As String

'флаги несуществования картинок
Public NoPicFrontFaceFlag As Boolean
Public NoPic1Flag As Boolean
Public NoPic2Flag As Boolean
Public NoPic3Flag As Boolean
Public NoPicActFlag As Boolean

Public DragCover As Boolean 'тянем ли текст обложки
Public ExitSVC As Boolean 'флаг завершения проги

Public LastLangFile As String 'уже загруженный файл локализации
Public NoSetColorFlag As Boolean 'true если цвета надо устрановить (поменять)

Public timerflag As Boolean 'что таймер отработал 1 раз

Public fnd As Integer 'файл для дебага

Public CoverWindTop As Long 'позиции формы FormShowPic
Public CoverWindLeft As Long
Public ScrShotWindTop As Long
Public ScrShotWindLeft As Long
Public IsCoverShowFlag As Boolean

Public SlideShowLastGoodPic As Integer 'последняя удачно показаный скриншот (пикча есть)
Public SlideShowLastFlag As Integer

Public isMPGflag As Boolean 'c чем имеем дело
Public isAVIflag As Boolean
Public isDShflag As Boolean 'какой другой DirectShow file (mov)

Public Const FW_NORMAL As Integer = 400   ' Normal font weight.
Public Const FW_BOLD As Integer = 700

Public NewDiskAddFlag As Boolean 'Добавляем новый диск - плюсовать CDN
Public OpenAddmovFlag As Boolean 'Добавляем новый файл (открыт ави)
Public AppendMovieFlag As Boolean 'Дополняем к открытому другой

Public dbMovieNameInd As Integer 'позиция поля в базе
Public dbLabelInd As Integer
Public dbGenreInd As Integer
Public dbYearInd As Integer
Public dbCountryInd As Integer
Public dbDirectorInd As Integer
Public dbActerInd As Integer
Public dbTimeInd As Integer
Public dbResolutionInd As Integer
Public dbAudioInd As Integer
Public dbFpsInd As Integer
Public dbFileLenInd As Integer
Public dbCDNInd As Integer
Public dbMediaTypeInd As Integer
Public dbVideoInd As Integer
Public dbSubTitleInd As Integer
Public dbLanguageInd As Integer
Public dbRatingInd As Integer
Public dbFileNameInd As Integer
Public dbDebtorInd As Integer
Public dbsnDiskInd As Integer
Public dbOtherInd As Integer
Public dbCoverPathInd As Integer
Public dbMovieURLInd As Integer
Public dbAnnotationInd As Integer

Public dbCheckedInd As Integer
Public dbSnapShot1Ind As Integer
Public dbSnapShot2Ind As Integer
Public dbSnapShot3Ind As Integer
Public dbFrontFaceInd As Integer
Public dbKeyInd As Integer


Public dbFirstField As String 'имя первого поля в базе

'ПО Listview
Public Const lvHeaderIndexPole As Integer = 25 ' 19 'номер поля индекса заголовка lv (lvIndexPole + 1название)
Public Const lvIndexPole As Integer = 24 ' 18 'номер поля индекса  - кол-во до аннотации в базе

Public Const LB_SETHORIZONTALEXTENT = &H194 'гориз. скролл листбокса
Public Const LB_SETSEL = &H185&
Public Const SRCCOPY = &HCC0020
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, _
   ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
   ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, _
   ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public FrmMainState As Integer 'хранит состояние главного окна

'input - звездочки
Private Declare Function FindWindowEx Lib "user32" Alias _
  "FindWindowExA" (ByVal hwnd1 As Long, ByVal hwnd2 As Long, _
   ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SetTimer& Lib "user32" _
  (ByVal hwnd&, ByVal nIDEvent&, ByVal uElapse&, ByVal _
   lpTimerFunc&)
Private Declare Function KillTimer& Lib "user32" _
  (ByVal hwnd&, ByVal nIDEvent&)
Const EM_SETPASSWORDCHAR = &HCC
Public Const NV_INPUTBOX As Long = &H5000&
'''
'серийник диска
Private Declare Function GetVolumeInformation _
    Lib "kernel32" Alias "GetVolumeInformationA" _
    (ByVal lpRootPathName As String, _
    ByVal pVolumeNameBuffer As String, _
    ByVal nVolumeNameSize As Long, _
    lpVolumeSerialNumber As Long, _
    lpMaximumComponentLength As Long, _
    lpFileSystemFlags As Long, _
    ByVal lpFileSystemNameBuffer As String, _
    ByVal nFileSystemNameSize As Long) As Long
'''''

'Public FindInAnnotFlag As Boolean ' для показа соответствующей аннотации при мультиселектном поиске в аннотации

'для смены языка для клипборда
Private Const KL_NAMELENGTH = 9
Private Declare Function GetKeyboardLayoutName Lib "user32" _
        Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Private Declare Function LoadKeyboardLayout Lib "user32" _
        Alias "LoadKeyboardLayoutA" (ByVal HKL As String, _
                ByVal Flags As Long) As Long
                
'for regional settings
Private Const LOCALE_SDECIMAL = &HE
'Private Const LOCALE_STHOUSAND = &HF
'Private Const WM_SETTINGCHANGE = &H1A
'Private Const HWND_BROADCAST = &HFFFF&
'Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Const LVM_FIRST As Integer = &H1000
    Public Const LVM_SETCOLUMNWIDTH As Integer = (LVM_FIRST + 30)
    Public Const LVSCW_AUTOSIZE As Integer = -1
    Public Const LVSCW_AUTOSIZE_USEHEADER As Integer = -2
Public Const LVM_GETSUBITEMRECT = LVM_FIRST + 56
Public Const LVM_SUBITEMHITTEST = LVM_FIRST + 57

Public Declare Function GetACP Lib "kernel32" () As Long
Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public SeparadorDecimal As String
'public Dim SeparadorMiles As String

Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
'Public Const HWND_TOP = 0
'Public Const HWND_BOTTOM = 1
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1

'+ комбики
Public Const SWP_NOSENDCHANGING = &H400

'for selcount
'Public Const LVM_FIRST = &H1000
Public Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)
'Public SelCount As Long
'
Public frmShow_xPos As Long 'позиции окна FormShowPic
Public frmShow_yPos As Long

'
Public NoDBFlag As Boolean 'открыта ли база

'Public ExportUpDownFlag As Boolean 'флаг для разворачивания,сворачивания настроек экспорта
Public SelRows() As Long ' заполн с 1 таблица хранения индексов помеченных в Листвью строк для импорта
Public CheckRows() As Long ' -..- чекнутых
Public MultiSel As Boolean 'помечено ли несколько строк
Public SelRowsKey() As String 'c 1 таблица хранения ключей помеченных в Листвью строк для импорта
Public CheckRowsKey() As String ' -..- чекнутых
'для примера сбора ключей выделенных строк LV см саб frmmain.DelMovies

Public FirstActivateFlag As Boolean 'флаг первой активации окна с Листвью

'aviinfo
'Public Type typCodecInf
'    FourCC As String * 4
'    Description As String
'End Type
'Public CodecList() As typCodecInf
'Public NumCodec As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'''''''''''''''''''''''''''''''''''''
Public Const OF_READ = &H0&
Public Declare Function lopen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Public Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
'Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function GetFileSizeEx Lib "kernel32.dll" _
 (ByVal hFile As Long, ByRef lpFileSize As Currency) _
 As Long


'Declare Function GetShortPathName Lib "KERNEL32" _
        Alias "GetShortPathNameA" (ByVal lpszLongPath As _
        String, ByVal lpszShortPath As String, ByVal _
        cchBuffer As Long) As Long
        
Public VideoStand As String 'PAL NTSC
Public Ratio As Single 'ширина на высоту для AVI , для DS MMI_Format
Public PixelRatio As Single '(118/81) - pal 16:9
Public PixelRatioSS As Single 'для маленьких окон скриншотов - сейчас = PicSS1.Height = PixelRatioSS
'16x9 pal = 1.459
'16x9 ntsc = 1.215
'4x3 pal = 1.094
'4x3 ntsc = 0.911
'Public MMI_Ratio As String '1.333 из mminfo это DERIVED AR
Public MMI_Format As Currency 'Single 'вычисленное 4/3 из mminfo
Public MMI_Format_str As String '4/3 из mminfo

'Public CDSerialCur As String 'серийник сд при открытии/редактировании
Public SameCDLabel As String 'если диск тот-же вписывать метку автоматом в редактор
Public CheckSameDisk As Boolean ' чекать ли что диск уже есть в базе (для авто?)

Public MediaSN As String 'запомнить серийник
Public MediaType As String 'запомнить тип носителя
Public IsSameCdFlag As Boolean 'новый фильм на этом же cd

Public LVSelectChanged As Boolean 'для статистики, что поменяли выделенные поля

'Public Position As String
Public Frames As Long
Public TimeL As Currency 'Long 'Long время мпега
Public TimesX100 As Currency 'Long 'Long 'TimeS * 100
Public lastRendedAVI As Long 'позиция
Public lastRendedMPG As Single

Public aviName As String 'имя файла при открытии в редакторе
Public mpgName As String
Public DShName As String '(mov etc)

Public SlideShowFlag As Integer


Public SaveCoverFlag As Boolean 'сохранять ли картинки (был ли апдейт)
Public SavePic1Flag As Boolean
Public SavePic2Flag As Boolean
Public SavePic3Flag As Boolean 'сохранять ли картинки (был ли апдейт)
Public SavePicActFlag As Boolean

'Public LastRSedit As Long 'какую запись редактируем в текущий момент

Public InitFlag As Boolean 'надо перечитать базу в LV
Public actInitFlag As Boolean
'Public NewInitFlag As Boolean 'добавлена новая запись

Public From_m_pGF As Long
Public AviWidth As Integer
Public AviHeight As Integer
Public ActNotManualClick As Boolean 'не кликать актера из set selected

Public LastIndAct As Long 'индекс строки у LVActer (полю индекса)
                           '(но использ как CurAct если LVActer.Sorted = False)
Public CurAct As Long 'индекс сортированной (помеченной) строки lvacter
Public CurActKey As String ' c "" текущий ключ lvActer
Public ToActFromLV As Long 'типа CurAct, куда переходить в актерах при клике на меню перехода на актера

Public ActFlag As Boolean ' флаг редактирование/добавления актера (Не при удалении) - обший наверно
Public ActEditFlag As Boolean 'флаг редактирования актера
Public ActNewFlag As Boolean 'флаг добавления актера


Public LastInd As Long 'индекс записи у LV равный текущему полю
'Public AllAccess As Boolean ' давать ли доступ на редактирование сетевой базы

'Public VFont As Font
'Public HFont As Font
Public VFontColor As Long
Public HFontColor As Long
Public LVFontColor As Long 'цвет текста списков и текстов
Public LVBackColor As Long 'цвет фона списков и текстов
Public LVHighLightLong As Long 'цвет выделения списков
Public CoverVertBackColor As Long 'цвет фона боковых надписей обложки
Public CoverHorBackColor As Long 'цвет фона аннотации обложки


Public lngFileName As String 'имя файла локализации

Public isWindowsNt As Boolean
Public bdname As String
Public abdname As String
Public LastVMI As Integer ' посл пункт верт меню
Public PrevVMI As Integer ' предп пункт верт меню (для ESC)

Public QJPG As Long
'Public CoverMoveFlag As Boolean 'подвигать обложку в редакторе
'Public TextVideoHid As String 'видео кодек
Public aferror As Boolean ' true если ошибка cAviFrameExtract
Public FindFilePath As String
Public BaseReadOnly As Boolean
Public BaseReadOnlyU As Boolean 'используется другими
Public BaseAReadOnly As Boolean
Public BaseAReadOnlyU As Boolean

Public FirstLVFill As Boolean ' первое заполнение LV

Public CurLVKey As String 'текущий индекс в листвью
Public CurSearch As Long 'текущий индекс в листвью c 1
'Public CurBasePoint As Long 'текущий индекс по базе =lastind


Public pos1 As Long, pos2 As Long, pos3 As Long ' фреймы текущих скриншотов

'Public BigIndAct As Long 'наибольший индекс(subitem(1)) при последнем заполнении LVActer

Public addflag As Boolean 'флаг, что создается новая запись, false - при первом же сейве
Public editFlag As Boolean
'Public delFlag As Boolean 'только что удаляли запись

Public PPMin As Long
Public PPMax As Long

Public Declare Function ShellExecute _
    Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

'' Constants - ShellExecute Return Codes (Errors)
'Public Const ERROR_OOM = 0               ' The operating system is out of memory or resources.
'Public Const ERROR_FILE_NOT_FOUND = 2    ' The specified file was not found.
'Public Const ERROR_PATH_NOT_FOUND = 3    ' The specified path was not found.
'Public Const ERROR_BAD_FORMAT = 11       ' The .exe file is invalid (non-Win32R .exe or error in .exe image).
'Public Const SE_ERR_ACCESSDENIED = 5     ' The operating system denied access to the specified file.
'Public Const SE_ERR_ASSOCINCOMPLETE = 27 ' The file name association is incomplete or invalid.
'Public Const SE_ERR_DDEBUSY = 30         ' The DDE transaction could not be completed because other DDE transactions were being processed.
'Public Const SE_ERR_DDEFAIL = 29         ' The DDE transaction failed.
'Public Const SE_ERR_DDETIMEOUT = 28      ' The DDE transaction could not be completed because the request timed out.
'Public Const SE_ERR_DLLNOTFOUND = 32     ' The specified dynamic-link library was not found.
'Public Const SE_ERR_FNF = 2              ' The specified file was not found.
'Public Const SE_ERR_NOASSOC = 31         ' There is no application associated with the given file name extension. This error will also be returned if you attempt to print a file that is not printable.
'Public Const SE_ERR_OOM = 8              ' There was not enough memory to complete the operation.
'Public Const SE_ERR_PNF = 3              ' The specified path was not found.
'Public Const SE_ERR_SHARE = 26           ' A sharing violation occurred.


''============================================
''nShowCmd Constants
''============================================
'Public Const SW_HIDE = 0
'Public Const SW_NORMAL = 1
'Public Const SW_SHOWMINIMIZED = 2
'Public Const SW_SHOWMAXIMIZED = 3
'Public Const SW_SHOWNOACTIVATE = 4
'Public Const SW_SHOW = 5
'Public Const SW_MINIMIZE = 6
'Public Const SW_SHOWMINNOACTIVE = 7
'Public Const SW_SHOWNA = 8
'Public Const SW_RESTORE = 9
'Public Const SW_MAX = 10

'Private Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As Long) As Long

'Private Const ERROR_FILE_NOT_FOUND = 2&
'Private Const ERROR_PATH_NOT_FOUND = 3&
'Private Const ERROR_BAD_FORMAT = 11&
'Private Const SE_ERR_ACCESSDENIED = 5        ' access denied
'Private Const SE_ERR_ASSOCINCOMPLETE = 27
'Private Const SE_ERR_DDEBUSY = 30
'Private Const SE_ERR_DDEFAIL = 29
'Private Const SE_ERR_DDETIMEOUT = 28
'Private Const SE_ERR_DLLNOTFOUND = 32
'Private Const SE_ERR_FNF = 2                ' file not found
'Private Const SE_ERR_NOASSOC = 31
'Private Const SE_ERR_PNF = 3                ' path not found
'Private Const SE_ERR_OOM = 8                ' out of memory
'Private Const SE_ERR_SHARE = 26

Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
'Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
    
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
'Public SecName As String
Public INIFILE As String 'для главной формы
Public iniFileName As String
Public iniGlobalFileName As String
Public CodecsFileName As String

'from tut
Public Declare Function AVIFileOpen Lib "avifil32.dll" (ByRef ppfile As Long, ByVal szFile As String, ByVal uMode As Long, ByVal pclsidHandler As Long) As Long  'HRESULT
'Global Const OF_READ  As Long = &H0
'Global Const OF_WRITE As Long = &H1
'#define OF_READWRITE        0x00000002
'#define OF_SHARE_COMPAT     0x00000000
'#define OF_SHARE_EXCLUSIVE  0x00000010
Global Const OF_SHARE_DENY_WRITE As Long = &H20
Global Const streamtypeVIDEO       As Long = 1935960438 'equivalent to: mmioStringToFOURCC("vids", 0&)
Public Declare Function AVIFileGetStream Lib "avifil32.dll" (ByVal pfile As Long, ByRef ppaviStream As Long, ByVal fccType As Long, ByVal lParam As Long) As Long
Public Declare Function AVIFileInfo Lib "avifil32.dll" (ByVal pfile As Long, pfi As AVI_FILE_INFO, ByVal lSize As Long) As Long 'HRESULT
Public Declare Function AVIStreamInfo Lib "avifil32.dll" (ByVal pAVIStream As Long, ByRef psi As AVI_STREAM_INFO, ByVal lSize As Long) As Long
Public Declare Function AVIStreamRelease Lib "avifil32.dll" (ByVal pavi As Long) As Long 'ULONG
Public Declare Function AVIFileRelease Lib "avifil32.dll" (ByVal pfile As Long) As Long
Global Const AVIERR_OK As Long = 0&
Public Declare Function AVIStreamStart Lib "avifil32.dll" (ByVal pavi As Long) As Long
Public Declare Function AVIStreamGetFrameOpen Lib "avifil32.dll" (ByVal pAVIStream As Long, _
                                                                ByRef bih As Any) As Long 'returns pointer to GETFRAME object on success (or NULL on error)
Public Declare Function AVIStreamGetFrame Lib "avifil32.dll" (ByVal pGetFrameObj As Long, _
                                                                ByVal lPos As Long) As Long 'returns pointer to packed DIB on success (or NULL on error)

Public Declare Function AVIStreamLength Lib "avifil32.dll" (ByVal pavi As Long) As Long

Public Type AVI_RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Public Type AVI_STREAM_INFO
    fccType As Long
    fccHandler As Long
    dwFlags As Long
    dwCaps As Long
    wPriority As Integer
    wLanguage As Integer
    dwScale As Long
    dwRate As Long
    dwStart As Long
    dwLength As Long
    dwInitialFrames As Long
    dwSuggestedBufferSize As Long
    dwQuality As Long
    dwSampleSize As Long
    rcFrame As AVI_RECT
    dwEditCount  As Long
    dwFormatChangeCount As Long
    szName As String * 64
End Type

Public Type AVI_FILE_INFO  '108 bytes?
    dwMaxBytesPerSecond As Long
    dwFlags As Long
    dwCaps As Long
    dwStreams As Long
    dwSuggestedBufferSize As Long
    dwWidth As Long
    dwHeight As Long
    dwScale As Long
    dwRate As Long
    dwLength As Long
    dwEditCount As Long
    szFileType As String * 64
End Type

'//BITMAP DEFINES (from mmsystem.h)
Public Type BITMAPINFOHEADER '40 bytes
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

'rotate text
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal w As Long, ByVal e As Long, ByVal o As Long, ByVal w As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal cp As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

' Data structure needed for Windows API call (GetSystemTime) для врем файлов
Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
' Windows API call to get system time from os
Public Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

'DebugMode
Declare Function GetModuleFileName Lib "kernel32" Alias _
"GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, _
ByVal nSize As Long) As Long

'иконки в меню
'Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
'Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
'Public Const MF_BITMAP = &H4&
'Type MENUITEMINFO
'cbSize As Long
'fMask As Long
'fType As Long
'fState As Long
'wID As Long
'hSubMenu As Long
'hbmpChecked As Long
'hbmpUnchecked As Long
'dwItemData As Long
'dwTypeData As String
'cch As Long
'End Type
'Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
'Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Boolean, lpMenuItemInfo As MENUITEMINFO) As Boolean
'Public Const MIIM_ID = &H2
'Public Const MIIM_TYPE = &H10
'Public Const MFT_STRING = &H0&

'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
'Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
        ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
        ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Private Const GWL_EXSTYLE = (-20)
'Private Const GWL_STYLE = (-16)
'Private Const ES_NUMBER = &H2000

'Private Const SWP_FRAMECHANGED = &H20
'Private Const SWP_NOMOVE = &H2
'Private Const SWP_NOOWNERZORDER = &H200
'Private Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4

Public Enum enWindowStyles
    WS_BORDER = &H800000
    'WS_CAPTION = &HC00000
    WS_CHILD = &H40000000
    'WS_CLIPCHILDREN = &H2000000
    'WS_CLIPSIBLINGS = &H4000000
    'WS_DISABLED = &H8000000
    WS_DLGFRAME = &H400000
    'WS_EX_ACCEPTFILES = &H10&
    'WS_EX_DLGMODALFRAME = &H1&
    'WS_EX_NOPARENTNOTIFY = &H4&
    'WS_EX_TOPMOST = &H8&
    'WS_EX_TRANSPARENT = &H20&
    'WS_EX_TOOLWINDOW = &H80&
    'WS_GROUP = &H20000
    'WS_HSCROLL = &H100000
    'WS_MAXIMIZE = &H1000000
    'WS_MAXIMIZEBOX = &H10000
    'WS_MINIMIZE = &H20000000
    'WS_MINIMIZEBOX = &H20000
    'WS_OVERLAPPED = &H0&
    'WS_POPUP = &H80000000
    'WS_SYSMENU = &H80000
    'WS_TABSTOP = &H10000
    WS_THICKFRAME = &H40000
    WS_VISIBLE = &H10000000
    'WS_VSCROLL = &H200000
    '\\ New from 95/NT4 onwards
    'WS_EX_MDICHILD = &H40
    WS_EX_WINDOWEDGE = &H100
    WS_EX_CLIENTEDGE = &H200
    'WS_EX_CONTEXTHELP = &H400
    'WS_EX_RIGHT = &H1000
    'WS_EX_LEFT = &H0
    'WS_EX_RTLREADING = &H2000
    'WS_EX_LTRREADING = &H0
    'WS_EX_LEFTSCROLLBAR = &H4000
    'WS_EX_RIGHTSCROLLBAR = &H0
    'WS_EX_CONTROLPARENT = &H10000
    WS_EX_STATICEDGE = &H20000
    'WS_EX_APPWINDOW = &H40000
    'WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
    'WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
End Enum

Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
'''''

'For Tokenize
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (dst As Any, src As Any, ByVal nBytes As Long)
Private Declare Sub RtlZeroMemory Lib "kernel32" (dst As Any, ByVal nBytes As Long)
Private Type SAFEARRAY1D
    cDims           As Integer
    fFeatures       As Integer
    cbElements      As Long
    cLocks          As Long
    pvData          As Long
    cElements       As Long
    lLbound         As Long
End Type

Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal olestr As Long, ByVal bLen As Long) As Long

Public Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Public Const SB_HORZ = 0

'for array sort
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByRef lpSource As Any, ByVal iLen As Long)
Public Enum SortOrder
   SortAscending = 0
   SortDescending = 1
End Enum

''''''''''для Array2Picture и наоборот
'STG.tlb
'Зарегистрировать и подключить к проекту
'regtlib STRM.tlb
'IStream interface ....

Private Type PictureHeader
   Magic As Long
   Size As Long
End Type
Private Declare Function CreateStreamOnHGlobal Lib "ole32" ( _
   ByVal hGlobal As Long, _
   ByVal fDeleteOnRelease As Long, _
   ppstm As IStream) As Long

Private Declare Function GetHGlobalFromStream Lib "ole32" ( _
  ByVal pstm As IStream, _
  phglobal As Long) As Long

Private Declare Function GlobalSize Lib "kernel32" ( _
  ByVal hMem As Long) As Long

'Private Declare Function GlobalLock Lib "kernel32" ( _
  ByVal hMem As Long) As Long

'Private Declare Function GlobalUnlock Lib "kernel32" ( _
  ByVal hMem As Long) As Long

'Private Declare Function GlobalAlloc Lib "kernel32" ( _
  ByVal wFlags As Long, _
  ByVal dwBytes As Long) As Long

Const S_OK = 0
'Const PictureID = &H746C&

' Global Memory Flags
'Const GMEM_MOVEABLE = &H2
'Const GMEM_ZEROINIT = &H40
'Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long


''''''''''''''''''''''''''''''''''''''''''''''''
''для определения доступна ли dll. А как зарегистрирована ли?
'Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
'Private Const MAX_MESSAGE_LENGTH = 512
''API declarations
'Private Declare Function GetLastError Lib "KERNEL32" () As Long
'Private Declare Function FormatMessage Lib "KERNEL32" Alias "FormatMessageA" ( _
'ByVal dwFlags As Long, _
'lpSource As Any, _
'ByVal dwMessageId As Long, _
'ByVal dwLanguageId As Long, _
'ByVal lpBuffer As String, _
'ByVal nSize As Long, _
'Arguments As Long) As Long
'Private Declare Function LoadLibrary Lib "KERNEL32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
'Private Declare Function FreeLibrary Lib "KERNEL32" (ByVal hLibModule As Long) As Long
''''''''''''''''''''''''''''''
'Function IsDLLAvailable(ByVal DllFilename As String) As Boolean
''зарегистрирована ли dll
'    Dim hModule As Long
'
'    hModule = LoadLibrary(DllFilename) 'attempt to load DLL
'    If hModule > 32 Then
'        FreeLibrary hModule 'decrement the DLL usage counter
'        IsDLLAvailable = True 'Return true
'    Else
'        IsDLLAvailable = False 'Return False
'    End If
'End Function

'Sub ResizeAllScrollbars(frm As Object)
'    Dim hsbHeight As Single
'    Dim vsbWidth As Single
'    Dim ctrl As Control
'
'    Const SM_CXVSCROLL = 2
'    Const SM_CYHSCROLL = 3
'
'    ' Determine suggested scrollbars' size (in twips)
'    hsbHeight = GetSystemMetrics(SM_CYHSCROLL) * Screen.TwipsPerPixelY
'    vsbWidth = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
'
'    ' iterate on all the controls on the form
'    For Each ctrl In frm.Controls
'    Debug.Print TypeName(ctrl)
'        Select Case TypeName(ctrl)
'            Case "HScrollBar"
'                ctrl.Height = hsbHeight
'            Case "VScrollBar"
'                ctrl.Width = vsbWidth
'            Case "FlatScrollBar"
'                If ctrl.Orientation = 1 Then
'                    ctrl.Height = hsbHeight
'                Else
'                    ctrl.Width = vsbWidth
'                End If
'        End Select
'    Next
'
'End Sub

'
'' Array2Picture
'' Converts a byte array (which contains a valid picture) to a
'' Picture object.
'Public Function Array2Picture(aBytes() As Byte) As StdPicture
'Dim oIPS As IPersistStream
'Dim oStream As IStream, hGlobal As Long, lPtr As Long
'Dim lSize As Long, hdr As PictureHeader
'Dim LRes As Long
'   ' Create a new empty
'   ' picture object
'   Set Array2Picture = New StdPicture
'
'   ' Get the IPersistStream interface
'   Set oIPS = Array2Picture
'
'   ' Calculate the array size
'   lSize = UBound(aBytes) - LBound(aBytes) + 1
'
'   ' Allocate global memory
'   hGlobal = GlobalAlloc(GHND, lSize + Len(hdr))
'
'   If hGlobal Then
'
'      ' Get a pointer to the memory
'      lPtr = GlobalLock(hGlobal)
'
'      ' Initialize the header
'      hdr.Magic = PictureID
'      hdr.Size = lSize
'
'      ' Write the header
'      MoveMemory ByVal lPtr, hdr, Len(hdr)
'
'      ' Copy the byte array to
'      ' the global memory
'      MoveMemory ByVal lPtr + Len(hdr), aBytes(0), lSize
'
'      ' Release the pointer
'      GlobalUnlock hGlobal
'
'      ' Create a IStream object
'      ' with the global memory
'      LRes = CreateStreamOnHGlobal(hGlobal, True, oStream)
'
'      If LRes = S_OK Then
'
'         ' Load the picture
'         ' from the stream
'         oIPS.Load oStream
'
'      End If
'
'      ' Release the IStream
'      ' object
'      Set oStream = Nothing
'
'   End If
'
'End Function

Public Function Picture2Array(ByVal oObj As StdPicture) As Byte()
Dim oIPS As IPersistStream
Dim oStream As IStream, hGlobal As Long, lPtr As Long
Dim lSize As Long, hdr As PictureHeader
Dim LRes As Long
Dim aBytes() As Byte

'If oObj = 0 Then Exit Function
On Error GoTo err

' Get the IPersistStream interface
Set oIPS = oObj

' Create a IStream object
' on global memory
LRes = CreateStreamOnHGlobal(0, True, oStream)

If LRes = S_OK Then
    ' Save the picture in the stream
    oIPS.Save oStream, True
    ' Get the global memory handle
    ' from the stream
    If GetHGlobalFromStream(oStream, hGlobal) = S_OK Then
        ' Get the memory size
        lSize = GlobalSize(hGlobal)
        ' Get a pointer to the memory
        lPtr = GlobalLock(hGlobal)
        If lPtr Then
            lSize = lSize - Len(hdr)
            ' Redim the array
            ReDim aBytes(0 To lSize - 1)
            ' Copy the data to the array
            MoveMemory aBytes(0), ByVal lPtr + Len(hdr), lSize
            Picture2Array = aBytes
        End If
        ' Release the pointer
        GlobalUnlock hGlobal
    End If
    ' Release the IStream
    ' object
    Set oStream = Nothing
End If

Exit Function
err:
ToDebug "Err_P2A: " & err.Description
Set oStream = Nothing
End Function


Public Function Tokenize04(Expression As String, ResultTokens() As String, Delimiters As String, Optional IncludeEmpty As Boolean, Optional PreserveIt As Boolean) As Long
' Tokenize02 by Donald, donald@xbeat.net
' modified by G.Beckmann, G.Beckmann@NikoCity.de
'возвращает кол-во токенов

'mzt     Const ARR_CHUNK& = 1024
    
    Dim cExp As Long, ubExpr As Long
    Dim cDel As Integer, ubDelim As Integer
    Dim aExpr() As Integer, aDelim() As Integer
    Dim sa1 As SAFEARRAY1D, sa2 As SAFEARRAY1D
    Dim cTokens As Long, iPos As Long
 
    ubExpr = Len(Expression)
    ubDelim = Len(Delimiters)
    
    sa1.cbElements = 2:     sa1.cElements = ubExpr
    sa1.cDims = 1:          sa1.pvData = StrPtr(Expression)
    RtlMoveMemory ByVal VarPtrArray(aExpr), VarPtr(sa1), 4
    
    sa2.cbElements = 2:     sa2.cElements = ubDelim
    sa2.cDims = 1:          sa2.pvData = StrPtr(Delimiters)
    RtlMoveMemory ByVal VarPtrArray(aDelim), VarPtr(sa2), 4
  
    If IncludeEmpty Then
        If PreserveIt Then
            ReDim Preserve ResultTokens(ubExpr)
        Else
            ReDim ResultTokens(ubExpr)
        End If
    Else
        If PreserveIt Then
            ReDim Preserve ResultTokens(ubExpr \ 2)
        Else
            ReDim ResultTokens(ubExpr \ 2)
        End If
    End If
    
    ubDelim = ubDelim - 1
    For cExp = 0 To ubExpr - 1
        For cDel = 0 To ubDelim
            If aExpr(cExp) = aDelim(cDel) Then
                If cExp > iPos Then
                    ResultTokens(cTokens) = Trim$(Mid$(Expression, iPos + 1, cExp - iPos))
                    cTokens = cTokens + 1
                ElseIf IncludeEmpty Then
                    ResultTokens(cTokens) = vbNullString
                    cTokens = cTokens + 1
                End If
                iPos = cExp + 1
                Exit For
            End If
        Next cDel
    Next cExp
  
    '/ remainder
    If (cExp > iPos) Or IncludeEmpty Then
        ResultTokens(cTokens) = Trim$(Mid$(Expression, iPos + 1))
        cTokens = cTokens + 1
    End If
  
    '/ erase or shrink
    If cTokens = 0 Then
        Erase ResultTokens()
    Else
        ReDim Preserve ResultTokens(cTokens - 1)
    End If
  
    '/ return ubound
    Tokenize04 = cTokens - 1
    
    '/ tidy up
    RtlZeroMemory ByVal VarPtrArray(aExpr), 4
    RtlZeroMemory ByVal VarPtrArray(aDelim), 4
End Function

Public Function DebugMode() As Boolean
Dim strFileName As String
Dim lngCount As Long
   
   strFileName = String(255, 0)
   lngCount = GetModuleFileName(App.hInstance, strFileName, 255)
   strFileName = left$(strFileName, lngCount)
   If UCase$(right$(strFileName, 7)) <> "VB6.EXE" Then
      DebugMode = False
   Else
      DebugMode = True
   End If
End Function

Function sTrimChars(s As String, c As String) As String
'эксклюзивчик
'чегото заменяем :) и обрезаем пробелы

Dim nPos As Integer, temp As String, temp2 As String

nPos = InStr(s, c)
Do While nPos > 0
 temp = left$(s, nPos - 1)
 temp2 = right$(s, Len(s) - Len(c) - Len(temp))
 If right$(temp, 1) = "," Then
  s = temp + " " + temp2
 Else
  s = temp + ", " + temp2
 End If
 nPos = InStr(s, c)
Loop
        
Do While right$(s, 2) = ", "
 s = left$(s, Len(s) - 2)
Loop
        
Do
 If right$(s, 1) = "," Or right$(s, 1) = ";" Or right$(s, 1) = "." Then
  s = left$(s, Len(s) - 1)
 Else
  Exit Do
 End If
Loop
    
sTrimChars = Trim$(s)
End Function
Function Change2lfcr(s As String) As String
'mzt Dim nPos As Integer, temp As String, temp2 As String
'меняем \ на перевод строки (ini msgbox)
If InStrB(s, "\") > 0 Then
Change2lfcr = Replace(s, "\", vbNewLine)
Else
Change2lfcr = s
End If
End Function
Public Sub ChTZ2ZPT(ByRef s As String)
If InStrB(1, s, ";") > 0 Then s = Replace(s, ";", ",")
End Sub

Public Function ReplaceCharacters(ByRef strText As String, _
    ByRef strUnwanted As String, ByRef strRepl As String) As String
Dim i As Integer
'mzt Dim ch As String
'strUnwanted - список заменяемых
For i = 1 To Len(strUnwanted)
 ' Replace the i-th unwanted character.
strText = Replace(strText, Mid$(strUnwanted, i, 1), strRepl)
Next

ReplaceCharacters = strText
End Function
Public Sub ReplaceFNStr(ByRef strData As String)
' замена символов для имен файлов

    'Replace invalid sings. \ / : * ? " < > |
    'strData = Replace$(strData, "_", " ", , , vbTextCompare)
    'strData = Replace$(strData, "ґ", "'", , , vbTextCompare)
    'strData = Replace$(strData, "`", "'", , , vbTextCompare)
    'strData = Replace$(strData, "{", "(", , , vbTextCompare)
    'strData = Replace$(strData, "[", "(", , , vbTextCompare)
    'strData = Replace$(strData, "]", ")", , , vbTextCompare)
    'strData = Replace$(strData, "}", ")", , , vbTextCompare)
    
    
    strData = Replace$(strData, "«", "(", , , vbTextCompare)
    strData = Replace$(strData, "»", ")", , , vbTextCompare)
    strData = Replace$(strData, "/", "_", , , vbTextCompare)
    strData = Replace$(strData, "\", "_", , , vbTextCompare)
    strData = Replace$(strData, ":", "-", , , vbTextCompare)
    strData = Replace$(strData, "*", "_", , , vbTextCompare)
    strData = Replace$(strData, "?", ".", , , vbTextCompare)
    strData = Replace$(strData, """", "_", , , vbTextCompare)
    strData = Replace$(strData, "<", "(", , , vbTextCompare)
    strData = Replace$(strData, ">", ")", , , vbTextCompare)
    strData = Replace$(strData, "|", "_", , , vbTextCompare)
    strData = Replace$(strData, "'", "`", , , vbTextCompare)
    
    strData = Replace$(strData, "–", "-", , , vbTextCompare) 'код 150 на 45
    
    
    'ReplaceFNStr = strData
End Sub
Public Sub ReplaceJSStr(ByRef strData As String)
'замена запрещенных символов для массива JavaScript
strData = Replace$(strData, """", "\""", , , vbTextCompare)

End Sub

Public Function Replace2Regional(strText As String) As String
Replace2Regional = strText
If InStrB(strText, ",") <> 0 Then Replace2Regional = Replace(strText, ",", SeparadorDecimal)
If InStrB(strText, ".") <> 0 Then Replace2Regional = Replace(strText, ".", SeparadorDecimal)
End Function
Public Function GetRNDFile(Optional PrependString As String = vbNullString, Optional Extension As String = vbNullString) As String
Dim lpSystemTime As SYSTEMTIME
Dim theTemp As String
Dim RandomNumber As Integer
     
'Get random number between 1 and 999 - will be appended to the result to ensure uniqueness
Randomize Timer
RandomNumber = Int((999 * Rnd) + 1)
     
GetSystemTime lpSystemTime
theTemp = PrependString & Trim$(lpSystemTime.wYear) & Format$(Trim$(lpSystemTime.wMonth), "00") & Format$(Trim$(lpSystemTime.wDay), "00") & Format$(Trim$(lpSystemTime.wHour), "00") & Format$(Trim$(lpSystemTime.wMinute), "00") & Format$(Trim$(lpSystemTime.wSecond), "00") & Format$(Trim$(lpSystemTime.wMilliseconds), "000") & Format$(Trim$(CStr(RandomNumber)), "000") & Extension
GetRNDFile = theTemp
 
End Function
 
Public Function GetNameFromPathAndName(ByVal sThePathAndName As String) As String
' Return a string containing the file name from a fully qualified file name. Larry
' If no path then return the file name anyhow
Dim sTemp1 As String                        'temporary string
Dim sTemp2 As String

sTemp1 = Trim$(sThePathAndName)             'save it
sTemp2 = GetPathFromPathAndName(sTemp1)     'get path
If InStrB(sThePathAndName, sTemp2) <> 0 Then
 sTemp1 = Mid$(sThePathAndName, Len(sTemp2) + 1)
End If
GetNameFromPathAndName = sTemp1             'now contains just file name
End Function

Public Function GetPathFromPathAndName(ByVal sThePathAndName As String) As String
' Return a string containing the file's path from a fully qualified file name. Larry
' Return "" if no path
Dim i As Integer                    'used in for/next loops
Dim sTemp As String

sTemp = Trim$(sThePathAndName)      'trim it
If InStrB(sTemp, "\") = 0 Then       'any backslash?
 Exit Function
End If
For i = Len(sTemp) To 1 Step -1     'find the right most one
 If Mid$(sTemp, i, 1) = "\" Then
  GetPathFromPathAndName = Mid$(sTemp, 1, i) 'now have just path
  Exit Function
 End If
Next
End Function

Public Function GetExtensionFromFileName(sTheFileAndExt As String, sTheFile As String) As String
'можно юзать getExtFromFile
' 1995/09/15 Return the file's extension from the filename and extension. Larry
' Return just the file name in sTheFile
Const csFrame = "+"                         'frame character
Const csPeriod = "."                        'ext follows this
Dim i As Integer
Dim iLen As Integer
'Dim iLoc As Integer                         'location of the period
Dim sTemp As String
Dim sFil As String                          'file
Dim sExt As String                          'extension
    
sTemp = csFrame & Trim$(sTheFileAndExt) & csFrame 'work with it here
If InStrB(sTheFileAndExt, ".") = 0 Then      'none, return what we found
 sTheFile = sTheFileAndExt
 Exit Function
End If

iLen = Len(sTemp)
For i = iLen To 1 Step -1                   'find the period
 If Mid$(sTemp, i, 1) = csPeriod Then
  'iLoc = i                            'got it
  sExt = Mid$(sTemp, i + 1)           'got the extension
  '2002/07/07 Change from 5 to 99
  If Len(sExt) > 1 And Len(sExt) < 99 Then 'OK, good
   sExt = Mid$(sExt, 1, Len(sExt) - 1) 'drop Frame character
  Else
   GoTo GetExtensionFromFileNameExit   'bad
  End If
  sFil = Mid$(sTemp, 1, i - 1)
  If Len(sFil) > 1 Then
   sFil = Mid$(sTemp, 2, i - 2)        'drop Frame character
  Else
   GoTo GetExtensionFromFileNameExit   'bad
  End If
  Exit For
 End If
Next

sTheFile = sFil                 'return what we found
GetExtensionFromFileName = sExt
Exit Function                   'bye
    
GetExtensionFromFileNameExit:
sTheFile = vbNullString                   'return blank
GetExtensionFromFileName = vbNullString   'not good
End Function


Public Function GetDecimalSymbol() As String
Dim Symbol As String
Dim iRet1 As Long
Dim iRet2 As Long
Dim lpLCDataVar As String
Dim pos As Integer
Dim Locale As Long
         
Locale = GetSystemDefaultLCID()
   
iRet1 = GetLocaleInfo(Locale, LOCALE_SDECIMAL, lpLCDataVar, 0)
Symbol = String$(iRet1, 0)
iRet2 = GetLocaleInfo(Locale, LOCALE_SDECIMAL, Symbol, iRet1)
pos = InStr(Symbol, vbNullChar) 'Chr$(0))
If pos > 0 Then
 Symbol = left$(Symbol, pos - 1)
 SeparadorDecimal = Symbol
End If

GetDecimalSymbol = SeparadorDecimal
End Function
Public Function getCurrLang() As String
Dim layoutname As String * KL_NAMELENGTH
Dim z As Long
    
z = GetKeyboardLayoutName(layoutname)
If z = 0 Then
 getCurrLang = vbNullString
Else
 getCurrLang = StrZ(layoutname)
End If
End Function
'Переключает на указанную sNewLang раскладку - возвращает старую раскладку
'am 030305_15:13:39
Public Function switchLang(sNewLang As String) As String
On Error Resume Next
'"00000419" - русская
'"00000409" - латинская
switchLang = getCurrLang
If StrComp(switchLang, sNewLang) <> 0 Then
 LoadKeyboardLayout sNewLang, 1
End If
End Function
'v_1.0.0 990630
Public Function StrZ(par As String) As String
Dim nSize As Long, i As Long 'mzt , rez As String
nSize = Len(par)
i = InStr(1, par, vbNullChar) - 1 'Chr(0)) - 1
If i > nSize Then i = nSize
If i < 0 Then i = nSize
StrZ = Mid$(par, 1, i)
End Function
Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime&)
'input - звездочки
Dim EditHwnd As Long
' CHANGE APP.TITLE TO YOUR INPUT BOX TITLE.
EditHwnd = FindWindowEx(FindWindow("#32770", App.title), 0, "Edit", "")
Call SendMessage(EditHwnd, EM_SETPASSWORDCHAR, Asc("*"), 0)
KillTimer hwnd, idEvent
End Sub

Public Sub ToDebug(s As String)
On Error Resume Next
FormDebug.TextDebug.Text = FormDebug.TextDebug.Text & Time & "> " & s & vbCrLf
FormDebug.TextDebug.SelStart = Len(FormDebug.TextDebug.Text)

If Not DebugFileFlagRW Then Exit Sub
Print #fnd, Time & "> " & s
End Sub


Public Function GetSerialNumber(ByVal sDrive As String, Optional ByRef VolumeName As String) As Long
'серийник диска

If Not Opt_GetVolumeInfo Then
    GetSerialNumber = 0
    VolumeName = vbNullString
    Exit Function
End If

If Len(sDrive) Then
    If InStr(sDrive, "\\") = 1 Then
        ' Make sure we end in backslash for UNC
        If right$(sDrive, 1) <> "\" Then
            sDrive = sDrive & "\"
        End If
    Else
        ' If not UNC, take first letter as drive
        sDrive = left$(sDrive, 1) & ":\"
    End If
Else
    ' Else just use current drive
    sDrive = vbNullString
End If

VolumeName = String$(64, 0)

' Grab S/N -- Most params can be NULL
Call GetVolumeInformation( _
     sDrive, VolumeName, Len(VolumeName), GetSerialNumber, ByVal 0&, ByVal 0&, vbNullString, 0)
VolumeName = TrimNull(VolumeName)
End Function


Public Sub Nuke(DirName As String)
'почикать папку с содержимым
Const ATTR_NORMAL = 0
Const ATTR_DIRECTORY = 16
Dim OriginalDir As String, filename As String

OriginalDir = CurDir$
ChDir DirName
filename = Dir$("*.*", ATTR_NORMAL)

Do While LenB(filename) <> 0
 Kill filename
 filename = Dir$
Loop

Do
 filename = Dir$("*.*", ATTR_DIRECTORY)
 While filename = "." Or filename = ".."
  filename = Dir$
 Wend
 If LenB(filename) = 0 Then Exit Do
 If Not exitNukeflag Then
  Nuke (filename)
 Else
  Exit Sub
 End If
Loop

ChDir OriginalDir
  
On Error Resume Next
RmDir DirName
If err.Number = 75 Then
 exitNukeflag = True
 ToDebug "В удаляемой папке кто-то сидит."
End If
End Sub

Public Function GetPCUserName() As String
Dim suser As String
Dim cnt As Long
Dim dl As Long

cnt = 199
suser = String$(200, 0)
dl = GetUserName(suser, cnt)

If dl <> 0 Then
 GetPCUserName = left$(suser, cnt - 1)
Else
 GetPCUserName = vbNullString
End If
End Function

'Public Function RemoveBorder(ByVal hWnd As Long)
'Dim lngRetVal As Long
'lngRetVal = GetWindowLong(hWnd, GWL_STYLE)
'lngRetVal = lngRetVal And (Not WS_BORDER) And (Not WS_DLGFRAME) And (Not WS_THICKFRAME)
'SetWindowLong hWnd, GWL_STYLE, lngRetVal
'lngRetVal = GetWindowLong(hWnd, GWL_EXSTYLE)
'lngRetVal = lngRetVal And (Not WS_EX_CLIENTEDGE) And (Not WS_EX_STATICEDGE) And (Not WS_EX_WINDOWEDGE)
'SetWindowLong hWnd, GWL_EXSTYLE, lngRetVal
'
'SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or _
'SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
'
'End Function

Public Function Quote(strVariable As String) As String
Quote = " ' " & strVariable & " ' "
End Function

Public Function MakeUserFile(fNm As String) As Boolean
Dim iFile As Integer

On Error GoTo ErrorHandler

iFile = FreeFile
Open fNm For Output As #iFile

Print #iFile, "#user.lng file for Sur Video Catalog."

'определить ветку по языку
Select Case LCase$(LastLanguage)
Case "русский"

Print #iFile, "[ExportPreset]"
Print #iFile, "Фильм мини=1,0,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
Print #iFile, "Фильм полный=1,0,1,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,1"
Print #iFile, ""
Print #iFile, "[Genre]"
Print #iFile, "1=Авантюрный"
Print #iFile, "2=Анимационный"
Print #iFile, "3=Боевик"
Print #iFile, "4=Вестерн"
Print #iFile, "5=Война"
Print #iFile, "6=Детектив"
Print #iFile, "7=Для взрослых"
Print #iFile, "8=Документальный"
Print #iFile, "9=Драма"
Print #iFile, "10=Исторический"
Print #iFile, "11=Комедия"
Print #iFile, "12=Криминал"
Print #iFile, "13=Мелодрама"
Print #iFile, "14=Мистика"
Print #iFile, "15=Музыкальный"
Print #iFile, "16=Мультипликация"
Print #iFile, "17=Пародия"
Print #iFile, "18=Приключения"
Print #iFile, "19=Притча"
Print #iFile, "20=Романтический"
Print #iFile, "21=Семейный"
Print #iFile, "22=Театр"
Print #iFile, "23=Триллер"
Print #iFile, "24=Ужасы"
Print #iFile, "25=Фантастика"
Print #iFile, "26=Философский"
Print #iFile, "27=Фэнтези"
Print #iFile, "28=Экшен"
Print #iFile, ""
Print #iFile, "[Country]"
Print #iFile, "1=Австралия"
Print #iFile, "2=Австрия"
Print #iFile, "3=Бельгия"
Print #iFile, "4=Бразилия"
Print #iFile, "5=Великобритания"
Print #iFile, "6=Германия"
Print #iFile, "7=Голландия"
Print #iFile, "8=Дания"
Print #iFile, "9=Италия"
Print #iFile, "10=Канада"
Print #iFile, "11=Китай"
Print #iFile, "12=Корея"
Print #iFile, "13=Мексика"
Print #iFile, "14=Польша"
Print #iFile, "15=Россия"
Print #iFile, "16=США"
Print #iFile, "17=Тайвань"
Print #iFile, "18=Франция"
Print #iFile, "19=Швейцария"
Print #iFile, "20=Швеция"
Print #iFile, "21=Япония"
Print #iFile, ""
Print #iFile, "[Site]"
Print #iFile, "1=http://dvd-film-shop.ru/search/index.html"
Print #iFile, "2=http://www.sharereactor.ru/cgi-bin/mzsearch.cgi"
Print #iFile, "3=http://www.videoguide.ru/"
Print #iFile, "4=http://www.rmdb.ru/"
Print #iFile, "5=http://www.kino-mir.ru/films"
Print #iFile, "6=http://www.movies.nnov.ru/"
Print #iFile, "7=http://www.kinopoisk.ru/level/7/"
Print #iFile, "8=http://www.zone5.ru/movies/"
Print #iFile, "9=http://www.kino.join.com.ua/"
Print #iFile, "10=http://www.kinomania.ru/"
Print #iFile, "11=http://www.kinokadr.ru/search/?query="
Print #iFile, "12=http://www.world-art.ru/"
Print #iFile, "13=http://www.ruscico.com/search.php?lang=ru"
Print #iFile, ""
Print #iFile, "[Language]"
Print #iFile, "1=русский дублированный"
Print #iFile, "2=русский многоголосый"
Print #iFile, "3=русский один голос"
Print #iFile, "4=русский два голоса"
Print #iFile, ""
Print #iFile, "[Subtitle]"
Print #iFile, "1=русский"
Print #iFile, "2=русский, английский"
Print #iFile, "3=английский"
Print #iFile, ""
Print #iFile, "[History]"

Case "english"

Print #iFile, "[ExportPreset]"
Print #iFile, "movie mini=1,0,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
Print #iFile, "movie full=1,0,1,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,1"
Print #iFile, ""
Print #iFile, "[Genre]"
Print #iFile, "1=Action"
Print #iFile, "2=Adult"
Print #iFile, "3=Adventure"
Print #iFile, "4=Animation"
Print #iFile, "5=Biblical"
Print #iFile, "6=Biography"
Print #iFile, "7=Comedy"
Print #iFile, "8=Crime"
Print #iFile, "9=Detective"
Print #iFile, "10=Documentary"
Print #iFile, "11=Drama"
Print #iFile, "12=Family"
Print #iFile, "13=Fantasy"
Print #iFile, "14=Horror"
Print #iFile, "15=Musical"
Print #iFile, "16=Mystery"
Print #iFile, "17=Sci-Fi"
Print #iFile, "18=Theater"
Print #iFile, "19=Triller"
Print #iFile, "20=War"
Print #iFile, "21=Western"
Print #iFile, "22=Melodrama"
Print #iFile, "23=Philosophical"
Print #iFile, "24=Romantic"
Print #iFile, "25=Parable"
Print #iFile, "26=Risky"
Print #iFile, ""
Print #iFile, "[Country]"
Print #iFile, "1=UK"
Print #iFile, "2=Germany"
Print #iFile, "3=Italy"
Print #iFile, "4=France"
Print #iFile, "5=Canada"
Print #iFile, "6=Russia"
Print #iFile, "7=USA"
Print #iFile, "8=Japan"
Print #iFile, "9=Australia"
Print #iFile, "10=Austria"
Print #iFile, "11=Belgium"
Print #iFile, "12=Brazil"
Print #iFile, "13=China"
Print #iFile, "14=Netherlands"
Print #iFile, "15=Poland"
Print #iFile, "16=Korea"
Print #iFile, "17=Sweden"
Print #iFile, "18=Switzerland"
Print #iFile, "19=Taiwan"
Print #iFile, "20=Mexico"
Print #iFile, ""
Print #iFile, "[Site]"
Print #iFile, "1=http://www.imdb.com/"
Print #iFile, "2=http://www.amazon.com/"
Print #iFile, "3=http://www.dvdempire.com/"
Print #iFile, ""
Print #iFile, "[Language]"
Print #iFile, "1=english"
Print #iFile, ""
Print #iFile, "[Subtitle]"
Print #iFile, "1=english"
Print #iFile, ""
Print #iFile, "[History]"

Case Else
Print #iFile, "[ExportPreset]"
Print #iFile, "[Country]"
Print #iFile, "[Site]"
Print #iFile, "[Language]"
Print #iFile, "[Subtitle]"
Print #iFile, "[History]"

End Select

'общие
Print #iFile, ""
Print #iFile, "[Media]"
Print #iFile, "1=CD"
Print #iFile, "2=CD-R"
Print #iFile, "3=CD-RW"
Print #iFile, "4=VCD"
Print #iFile, "5=SVCD"
Print #iFile, "6=DVD-Video"
Print #iFile, "7=DVD 5"
Print #iFile, "8=DVD 9"
Print #iFile, "9=DVD-R"
Print #iFile, "10=DVD+R"
Print #iFile, "11=DVD-RW"
Print #iFile, "12=DVD+RW"
Print #iFile, "13=Mini CD-R"
Print #iFile, "15=HD DVD"
Print #iFile, "16=Blu-Ray"
Print #iFile, "17=HDD"
Print #iFile, "18=CD-ROM"
Print #iFile, ""
Print #iFile, "[Comments]"
Print #iFile, "1=DVDrip"
Print #iFile, "2=Telesync (TS)"
Print #iFile, "3=Screener (SCR)"
Print #iFile, "4=DVDScreener (SCR)"
Print #iFile, "5=Workprint (WP)"
Print #iFile, "6=Telecyne (TC)"
Print #iFile, "7=VHSrip"
Print #iFile, "8=TVrip"
Print #iFile, "9=SATrip"
Print #iFile, "10=CamRip (CAM)"
Print #iFile, "11=LDrip"
Print #iFile, "12=VHSScr"
Print #iFile, "13=Dubbed"
Print #iFile, "14=Line.Dubbed"
Print #iFile, "15=Mic.Dubbed"
Print #iFile, "16=StraitToVideo (STV)"

MakeUserFile = True
Close #iFile
Exit Function

ErrorHandler:
ToDebug "Err_musINI: " & err.Description
End Function
Public Function MakeINI(fNm As String) As Boolean

Dim iFile As Long
Dim arr() As Variant
Dim i As Integer
Dim WFD As WIN32_FIND_DATA
Dim ret As Long
'Dim tmp As String

On Error GoTo ErrorHandler

iFile = FreeFile
Open App.Path & "\" & fNm For Output As #iFile

If LCase$(fNm) = "global.ini" Then

    Print #iFile, "[CD]"
    Print #iFile, "#D:\;C:\Video"
    Print #iFile, "CDdrive=" & GetFirstOptoDrive
    Print #iFile, ""
    Print #iFile, "[LIST]"
    Print #iFile, "LVWidth%=50"
    Print #iFile, "TVWidth=2500"
    Print #iFile, "ScrShotWidth%=40"
    Print #iFile, "LastItem=1"
    arr = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 10, 12, 11, 20, 17, 16, 9, 14, 13, 15, 19, 24, 18, 23, 21, 22, 25)
    For i = 1 To 25
        Print #iFile, "P" & i & "=" & arr(i)
    Next i
    Print #iFile, ""
    Print #iFile, "[EXPORT]"
    Print #iFile, "L0=True"
    Print #iFile, ""
    Print #iFile, "[COVER]"
    Print #iFile, ""
    Print #iFile, "[FONT]"
    Print #iFile, "VFontName=Arial"
    Print #iFile, "VFontSize=12"
    Print #iFile, "VFontBold=True"
    Print #iFile, "VFontItalic=False"
    Print #iFile, "HFontName=MS Sans Serif"
    Print #iFile, "HFontSize=9.75"
    Print #iFile, "HFontBold=False"
    Print #iFile, "HFontItalic=False"
    Print #iFile, "LVFontName=MS Sans Serif"
    Print #iFile, "LVFontSize=9.75"
    Print #iFile, "LVFontBold=False"
    Print #iFile, "LVFontItalic=False"
    Print #iFile, "VFontColor=0"
    Print #iFile, "HFontColor=0"
    Print #iFile, "LVFontColor=0"
    Print #iFile, "LVBackColor=15000275"
    Print #iFile, "CoverHorBackColor=16777215"
    Print #iFile, "CoverVertBackColor=16777215"
    Print #iFile, ""
    Print #iFile, "[LANGUAGE]"
    Print #iFile, "#Change LCount and add more language"
    Print #iFile, "lcount=2"
    Print #iFile, "#LastLang=English"
    'Print #iFile, "LastLang=Русский" 'выберем автоматом
    Print #iFile, "L1=Русский"
    Print #iFile, "L1File=rus.lng"
    Print #iFile, "L2=English"
    Print #iFile, "L2File=eng.lng"
    Print #iFile, "L3="
    Print #iFile, "L3File="
    Print #iFile, ""
    Print #iFile, "[GLOBAL]"
    ret = FindFirstFile(App.Path & "\firstrun.mdb", WFD)
    If ret > -1 Then
        Print #iFile, "BDCount=1"
        Print #iFile, "BDName1=" & App.Path & "\firstrun.mdb"
    Else
        Print #iFile, "BDCount=0"
    End If

Else    'взять текущие значения (как у активной базы, или у глобала)
'не передал сортировку

    Print #iFile, "[CD]"
    Print #iFile, "#D:\;C:\Video"
    Print #iFile, "CDdrive=" & ComboCDHid_Text
    Print #iFile, ""
    Print #iFile, "[LIST]"
    Print #iFile, "LVWidth%=" & LVWidth
    Print #iFile, "TVWidth=" & TVWidth
    Print #iFile, "ScrShotWidth%=" & SplitLVD
    Print #iFile, "LastItem=1"
    For i = 1 To 25
        Print #iFile, "C" & i & "=" & FrmMain.ListView.ColumnHeaders(i).Width
    Next i
    For i = 1 To 25
        Print #iFile, "P" & i & "=" & FrmMain.ListView.ColumnHeaders(i).Position
    Next i
    Print #iFile, "LVGrid=" & Opt_ShowLVGrid
    Print #iFile, "ColorDebt=" & Opt_Debtors_Colorize
    Print #iFile, ""
                        
    Print #iFile, "[EXPORT]"
    For i = 0 To 24
        Print #iFile, "L" & i & "=" & LstExport_Arr(i)
    Next i
    Print #iFile, "NumsOnPage=" & TxtNnOnPage_Text
    Print #iFile, "ExportDelimiter=""" & ExportDelim & """"
    For i = 0 To 2
        If Opt_HtmlJpgName = i Then
            Print #iFile, "OptHtml" & i & ".Caption=True"
        End If
    Next i
    Print #iFile, "Template=" & CurrentHtmlTemplate
    
    Print #iFile, "UseSubFolders=" & Opt_ExpUseFolders
    Print #iFile, "SubFolder1=" & Opt_ExpFolder1
    Print #iFile, "SubFolder2=" & Opt_ExpFolder2
    Print #iFile, "SubFolder3=" & Opt_ExpFolder3

    
    Print #iFile, ""
    Print #iFile, "[COVER]"
    Print #iFile, "txt_Stan_L=" & cov_stan.l
    Print #iFile, "txt_Stan_T=" & cov_stan.t
    Print #iFile, "txt_Stan_W=" & cov_stan.w
    Print #iFile, "txt_Stan_H=" & cov_stan.H
    Print #iFile, "txt_Conv_L=" & cov_conv.l
    Print #iFile, "txt_Conv_T=" & cov_conv.t
    Print #iFile, "txt_Conv_W=" & cov_conv.w
    Print #iFile, "txt_Conv_H=" & cov_conv.H
    Print #iFile, "txt_Dvd_L=" & cov_dvd.l
    Print #iFile, "txt_Dvd_T=" & cov_dvd.t
    Print #iFile, "txt_Dvd_W=" & cov_dvd.w
    Print #iFile, "txt_Dvd_H=" & cov_dvd.H
    Print #iFile, "txt_List_L=" & cov_list.l
    Print #iFile, "txt_List_T=" & cov_list.t
    Print #iFile, "txt_List_W=" & cov_list.w
    Print #iFile, "txt_List_H=" & cov_list.H
    Print #iFile, "ShowColNames=" & Opt_ShowColNames

    Print #iFile, ""
    Print #iFile, "[FONT]"
    Print #iFile, "VFontName=" & FontVert.name
    Print #iFile, "VFontSize=" & FontVert.Size
    Print #iFile, "VFontBold=" & FontVert.Bold
    Print #iFile, "VFontItalic=" & FontVert.Italic
    Print #iFile, "HFontName=" & FontHor.name
    Print #iFile, "HFontSize=" & FontHor.Size
    Print #iFile, "HFontBold=" & FontHor.Bold
    Print #iFile, "HFontItalic=" & FontHor.Italic
    Print #iFile, "LVFontName=" & FontListView.name
    Print #iFile, "LVFontSize=" & FontListView.Size
    Print #iFile, "LVFontBold=" & FontListView.Bold
    Print #iFile, "LVFontItalic=" & FontListView.Italic
    Print #iFile, "VFontColor=" & VFontColor
    Print #iFile, "HFontColor=" & HFontColor
    Print #iFile, "LVFontColor=" & LVFontColor
    Print #iFile, "LVBackColor=" & LVBackColor
    Print #iFile, "CoverHorBackColor=" & CoverHorBackColor
    Print #iFile, "CoverVertBackColor=" & CoverVertBackColor
    Print #iFile, "LVHighLight=" & LVHighLightLong

    Print #iFile, "VMcolor=" & VMSameColor
    Print #iFile, "StripedLV=" & StripedLV
    Print #iFile, "NoLVSelFrame=" & NoLVSelFrame
 
    Print #iFile, ""
    Print #iFile, "[GLOBAL]"
    Print #iFile, "ListAndInfo=" & Opt_UCLV_Vis
    Print #iFile, "FreeDVDFilters=" & Opt_UseOurMpegFilters
    Print #iFile, "SlideShowWindow=" & Opt_NoSlideShow
    Print #iFile, "GroupWindow=" & Opt_Group_Vis
    Print #iFile, "CenterShowPic=" & Opt_CenterShowPic
    Print #iFile, "QJPG=" & QJPG
    Print #iFile, "SaveBigPix=" & Opt_PicRealRes
    Print #iFile, "LVLoadOnlyTitle=" & Opt_LoadOnlyTitles
    Print #iFile, "SaveOptOnExit=" & Opt_AutoSaveOpt
    Print #iFile, "UseAspect=" & Opt_UseAspect
    Print #iFile, "SortOnStart=" & Opt_SortOnStart
    Print #iFile, "LoanAllSameLabels=" & Opt_LoanAllSameLabels
    
    Print #iFile, "LVEDIT=" & Opt_LVEDIT
    Print #iFile, "SaveFileWithPath=" & Opt_FileWithPath
    
    Print #iFile, "SortLVAfterEdit=" & Opt_SortLVAfterEdit
    Print #iFile, "SortLabelAsNum=" & Opt_SortLabelAsNum
    Print #iFile, "PutOtherInAnnot=" & Opt_PutOtherInAnnot

    
    Print #iFile, ""
    Print #iFile, "[AutoAdd]"
    Print #iFile, "chSubFolders=" & ch_chSubFolders
    Print #iFile, "chAvi=" & ch_chAviHid
    Print #iFile, "chDS=" & ch_chDSHid
    Print #iFile, "chShots=" & ch_chShots
    Print #iFile, "chNoMess=" & AutoNoMessFlag
    Print #iFile, "cAutoClose=" & ch_cAutoClose
    Print #iFile, "cEjectMedia=" & ch_cEjectMedia
    Print #iFile, "AVIsExt=" & extAvi
    Print #iFile, "DirectShowExt=" & extDS

End If

MakeINI = True
Close #iFile
Exit Function

ErrorHandler:
Close #iFile
ToDebug "mINI: " & err.Description
End Function
Public Sub KeepFormOnScreen(frm As Form)
' Keep 'frm' on the Screen.

On Error GoTo ErrorHandler

' If the form is off the screen then attempt to make it fit
' on the screen somehow.
If frm.Width <= Screen.Width Then
    If (frm.left + frm.Width) > Screen.Width Then
        frm.left = Screen.Width - frm.Width
    End If
    If frm.left < 0 Then
        frm.left = 0
    End If
End If

If frm.Height <= Screen.Height Then
    If (frm.top + frm.Height) > Screen.Height Then
        frm.top = Screen.Height - frm.Height - 100
    End If
    If frm.top < 0 Then
        frm.top = 0
    End If
End If

MyExit:

Exit Sub

ErrorHandler:
'MsgBox err.Description, vbExclamation
Resume MyExit

End Sub    'KeepFormOnScreen'


' Return the next word from this string. Remove
' the word from the string.
Public Function GetWord(txt As String) As String
Dim pos As Integer
'жрет двойные и более пробелы

    txt = Trim$(txt)
    pos = InStr(txt, " ")
    If pos < 1 Then
        GetWord = txt
        txt = vbNullString
    Else
        GetWord = left$(txt, pos - 1)
        txt = Trim$(right$(txt, Len(txt) - pos))
    End If
End Function

'Public Sub SetWrkSize()
'' Запрещаем VB кэшировать срань разную! (только для w2k)
''If Not OS_Version.dwPlatformID = &H2 Or OS_Version.dwMajorVersion < &H5 Then Exit Sub
'    If Not isWindowsNt Then Exit Sub
'    Call SetProcessWorkingSetSize(GetCurrentProcess, &HFFFF, &HFFFF)
'End Sub

Public Function pSaveDialog(Optional wndh As Long, Optional DTitle As String, Optional myFileName As String) As String
'диалог записи картинки в файл

Dim cd As cCommonDialog
Dim sfile As String
Dim fTitle As String
   
If wndh = 0 Then wndh = FrmMain.hwnd

Set cd = New cCommonDialog

sfile = myFileName 'дать начальное имя файлу

DoEvents 'last
'GIF (*.gif)|*.gif|
   If (cd.VBGetSaveFileName( _
      sfile, _
      fTitle, _
      Filter:="BMP (*.bmp)|*.bmp|JPG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF Uncompressed (*.tif)|*.tif", _
      FilterIndex:=2, _
      DlgTitle:=DTitle, _
      DefaultExt:="jpg", _
      Owner:=wndh)) Then
      pSaveDialog = sfile
      'fTitle - только имя файла без пути
   End If

'Set cd = Nothing
End Function
Public Function pSaveDialogBMP(Optional wndh As Long, Optional DTitle As String) As String
'диалог записи картинки только в BMP файл (movie)
Dim cd As cCommonDialog
Dim sfile As String
If wndh = 0 Then wndh = FrmMain.hwnd

Set cd = New cCommonDialog

DoEvents 'last

   If (cd.VBGetSaveFileName( _
      sfile, _
      Filter:="BMP (*.bmp)|*.bmp", _
      FilterIndex:=1, _
      DlgTitle:=DTitle, _
      DefaultExt:="bmp", _
      Owner:=wndh)) Then
      pSaveDialogBMP = sfile
   End If

'Set cd = Nothing
End Function
Public Function pLoadDialog(Optional DTitle As String, Optional filename As String) As String
'Dim cd As New cCommonDialog
Dim cd As cCommonDialog
Dim sfile As String
Dim indir As String

'если есть путь
Dim afn As String
Dim n As Long

On Error Resume Next

If editFlag Then    'впихнуть путь
    afn = FrmMain.ListView.SelectedItem.SubItems(dbFileNameInd)
    If Len(afn) <> 0 Then    'есть файл/ы
        n = InStr(afn, "|")
        If n > 0 Then afn = left$(afn, n - 1)    'первый файл
        indir = GetPathFromPathAndName(afn)
    End If
End If

DoEvents
Set cd = New cCommonDialog

sfile = filename

If (cd.VBGetOpenFileName( _
    sfile, _
    ReadOnly:=True, _
    HideReadOnly:=True, _
    Filter:="Multimedia files |*.avi;*.vid;*.divx;*.mpg;*.mpeg;*.vob;*.mov;*.asf;*.wmv;*.flv;*.mkv;*.mp4;*.3gp|All Files (*.*)|*.*", _
    FilterIndex:=1, _
    InitDir:=indir, _
    DlgTitle:=DTitle, _
    DefaultExt:="", _
    Owner:=frmEditor.hwnd)) Then
    pLoadDialog = sfile
End If

If Len(sfile) <> 0 Then LastMovieFolder = GetPathFromPathAndName(sfile)

Set cd = Nothing
End Function
Public Function pLoadPixDialog() As String
'Dim cd As New cCommonDialog
Dim cd As cCommonDialog
Dim sfile As String
   
'DoEvents
'Filter:="MS Access|*.mdb;*.amm|XML files|*.xml|All Files (*.*)|*.*|All Supported |*.mdb;*.xml;*.amm",

Set cd = New cCommonDialog

   If (cd.VBGetOpenFileName( _
      sfile, _
      Filter:="Images (BMP, GIF, JPG, PNG, TIF)|*.jpg;*.gif;*.png;*.tif;*.bmp|All files (*.*)| *.*", _
      FilterIndex:=1, _
      DefaultExt:="jpg", _
      Owner:=frmEditor.hwnd)) Then
      pLoadPixDialog = sfile
   End If

'FrmMain.Form_Resize
'FrmMain.SetFocus
'FrmMain.Refresh
'
Set cd = Nothing
End Function

Public Function GetFileNameFromEditor() As String
'найти в редакторе имя файла сложить с LastMovieFolder
'не делать, если файлов несколько
Dim sFN As String


sFN = frmEditor.TextFileName
If Len(sFN) = 0 Then Exit Function
If InStr(sFN, "|") > 0 Then Exit Function

If FileExists(sFN) Then ''так можно без пути остаться, если путь закеширован, то файл найдется и передастся без пути
'поэтому проверим есть ли путь
If GetNameFromPathAndName(sFN) <> sFN Then
    GetFileNameFromEditor = sFN
    Exit Function
End If
End If

If Len(LastMovieFolder) = 0 Then Exit Function
If FileExists(LastMovieFolder & sFN) Then
    GetFileNameFromEditor = LastMovieFolder & sFN
End If

End Function
Public Sub Pic2JPG(lImage As PictureBox, DBind As Integer, Field As String)
Dim img As ImageFile ', Pic As ImageFile
Dim IP As ImageProcess
Dim vec As Vector
'Dim stype As String

On Error GoTo wiaerr

Set vec = New Vector
Set IP = New ImageProcess

If lImage.Picture = 0 Then Exit Sub
vec.BinaryData = Picture2Array(lImage.Picture)

Set img = vec.ImageFile
Set vec = Nothing

'sType = wiaFormatBMP
'stype = wiaFormatJPEG
'stype = wiaFormatGIF
'stype = wiaFormatPNG
'sType = wiaFormatTIFF
'
'If Not img Is Nothing Then
'     Set p.Picture = img.ARGBData.Picture(lImage.Width, lImage.Height)
'End If

While (IP.Filters.Count > 0)
 IP.Filters.Remove 1
Wend
        
    IP.Filters.Add IP.FilterInfos("Convert").FilterID
    IP.Filters(1).Properties(1).Value = wiaFormatJPEG 'stype
    IP.Filters(1).Properties(2).Value = QJPG
    Set img = IP.Apply(img)
    Set IP = Nothing
    
If img Is Nothing Then
    'не вышло
    ToDebug "Err_2jpg: " & lImage.Width & "x" & lImage.Height
    Else
    
    Select Case DBind
        Case 1    'data
            rs.Fields(Field) = img.FileData.BinaryData
        Case 2    'ars
            ars.Fields(Field) = img.FileData.BinaryData
        Case 3    'dragdrop
            rsdd.Fields(Field) = img.FileData.BinaryData
    End Select

End If

Set img = Nothing
Exit Sub

'On Error GoTo wiaerr
wiaerr:
ToDebug "Err_p2j_WIA: " & err.Description
'MsgBox err.Description, vbCritical

End Sub


Public Function SearchCBO(ByRef cboCtl As ComboBox, ByVal sSearchCriteria As String, Optional ByVal bLIKE As Boolean = True) As Long
'<RR 08/08/2003 - VB/OUTLOOK GURU>
On Error GoTo No_Bugs

If bLIKE = False Then     'EXACT MATCH
    SearchCBO = SendMessageStr(cboCtl.hwnd, CB_FINDSTRINGEXACT, 0&, ByVal sSearchCriteria)
Else     'LIKE MATCH
    SearchCBO = SendMessageStr(cboCtl.hwnd, CB_FINDSTRING, 0&, ByVal sSearchCriteria)
End If

Exit Function
No_Bugs:
ToDebug "Err_SeaCBO:" & err.Description
End Function

Public Function SearchListBox(ByRef lbCtl As ListBox, ByVal sSearchCriteria As String, Optional ByVal bLIKE As Boolean = True) As Long
'возвратит index или -1
'On Error Resume Next
If bLIKE = False Then    'EXACT MATCH
    SearchListBox = SendMessage(lbCtl.hwnd, LB_FINDSTRINGEXACT, -1, ByVal sSearchCriteria)
Else
    SearchListBox = SendMessage(lbCtl.hwnd, LB_FINDSTRING, -1, ByVal sSearchCriteria)
End If
End Function

Public Function CalcFormat(s As String, Asp As String) As Currency    'Single
'вичмслим строку формата (4/3 16/9)
'Asp - строка представление числового значения аспекта из мминфо
's - дб с точкой для val
If Not isMPGflag Then
    If InStr(s, "/") Then
        Asp = Replace2Regional(Asp)
        CalcFormat = CCur(Asp)
        If CalcFormat = 0 Then CalcFormat = 4 / 3
    Else
        CalcFormat = Val(s)
    End If
    
    Exit Function
    
Else
    CalcFormat = 1.33333    'по умолчанию 4/3
End If

If InStr(s, "4/3") > 0 Then CalcFormat = 4 / 3: Exit Function
If InStr(s, "16/9") > 0 Then CalcFormat = 16 / 9: Exit Function
If InStr(s, "1.85") > 0 Then CalcFormat = 1.85: Exit Function
If InStr(s, "2.25") > 0 Then CalcFormat = 2.25: Exit Function
If InStr(s, "2.35") > 0 Then CalcFormat = 2.35: Exit Function
If InStr(s, "2.2") > 0 Then CalcFormat = 2.2: Exit Function 'ниже

'On Error Resume Next
'CalcFormat = Val(s) 'а то бывает 8, 12

End Function

Public Function ReadLang(Itm As String, Optional ss As String) As String
ReadLang = VBGetPrivateProfileString("Language", Itm, lngFileName, ss)
End Function
Public Function ReadLangStat(Itm As String, Optional ss As String) As String
ReadLangStat = VBGetPrivateProfileString("STATISTICS", Itm, lngFileName, ss)
End Function
Public Function ReadLangSR(Itm As String, Optional ss As String) As String
ReadLangSR = VBGetPrivateProfileString("SEARCH_REPLACE", Itm, lngFileName, ss)
End Function
Public Function ReadLangOpt(Itm As String, Optional ss As String) As String
ReadLangOpt = VBGetPrivateProfileString("OPTIONS", Itm, lngFileName, ss)
End Function
Public Function ReadLangFilt(Itm As String, Optional ss As String) As String
ReadLangFilt = VBGetPrivateProfileString("FILTER", Itm, lngFileName, ss)
End Function
Public Function ReadLangActFilt(Itm As String, Optional ss As String) As String
ReadLangActFilt = VBGetPrivateProfileString("ACTFILTER", Itm, lngFileName, ss)
End Function
Public Function ReadLangAuto(Itm As String, Optional ss As String) As String
ReadLangAuto = VBGetPrivateProfileString("AUTOADD", Itm, lngFileName, ss)
End Function


Public Sub FillTemplateCombo(ByVal fPath As String, Comb As VB.ComboBox)
Dim File_Name As String
'Dim tmpi As Integer
'tmpi = InStrRev(File_Name, ".")

If right$(fPath, 1) <> "\" Then fPath = fPath & "\"
   
File_Name = Dir$(fPath, vbDirectory)
Comb.Clear
Do While File_Name <> ""
 'tmpi = 0
 If File_Name <> "." And File_Name <> ".." Then
 
  'template
  If Comb.name = "CombTemplate" Then
  If LCase$(left$(File_Name, 4)) = "svc_" Then
   If InStr(1, File_Name, ".htm", vbTextCompare) Then
    Comb.AddItem File_Name
   End If
  End If
  End If
  
  'script
  If Comb.name = "ComboInfoSites" Then
   If InStr(1, File_Name, ".vbs", vbTextCompare) Then
    Comb.AddItem File_Name
   End If
  End If
        
  File_Name = Dir$
 Else
  File_Name = Dir$
 End If

Loop
 
'If Comb.ListCount > 0 Then Comb.Text = Comb.List(0)
End Sub

Public Function CheckNoNullValMyRS(ByRef F As Integer, ByRef rs2proc As DAO.Recordset) As String
'принимает номер поля базы
'работаем  с полученным rs2proc
If Not IsNull(rs2proc(F)) Then
    CheckNoNullValMyRS = rs2proc(F)
Else
    CheckNoNullValMyRS = vbNullString
End If
End Function

Public Function CheckNoNullVal(ByRef F As Integer) As String
'принимает номер поля базы
'работаем только с текущим rs
If Not IsNull(rs(F)) Then
    CheckNoNullVal = rs(F)
Else
    CheckNoNullVal = vbNullString
End If
End Function

Public Sub pTurnOffFullDrag()
   Dim lR As Long
   ' Get the full-drag state:
   If Not (SystemParametersInfo(SPI_GETDRAGFULLWINDOWS, 0&, lR, 0) = 0) Then
      If Not lR = 0 Then
         ' Store the fact we are changing:
         m_bSwitchOff = True
         ' Set the full-drag state:
         lR = SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, 0&, ByVal 0&, SPIF_SENDWININICHANGE)
      End If
   End If
End Sub

Public Sub pResetFullDrag()
'   If m_bSwitchOff Then
      SystemParametersInfo SPI_SETDRAGFULLWINDOWS, 1&, ByVal 0&, SPIF_SENDWININICHANGE
      m_bSwitchOff = False
'   End If
End Sub

Public Function CountryLocal(s As String) As String
'преобразование стран по iso (для мпег2)
Dim tmp As String, sCod As String
Dim i As Integer
Dim lcod() As String 'коды языков из строки
'vstrip выдает в языке ССru - СС спецсимволы - убрать

If Tokenize04(s, lcod(), ",;", False) > -1 Then

For i = 0 To UBound(lcod)
sCod = VBGetPrivateProfileString("ISO639", Trim$(right$(lcod(i), 2)), lngFileName)
If Len(sCod) = 0 Then sCod = Trim$(right$(lcod(i), 2))
tmp = tmp & sCod & ", "
Next i
CountryLocal = left$(tmp, Len(tmp) - 2)
End If
End Function

Public Function Time2sec(t As String) As Long
't hh:mm:ss - > ms
'обратная функция FormatTime(TextTimeMSHid)
Dim b() As String
Dim temp As Long

On Error GoTo err

    Const OneMinute As Long = 60
    Const OneHour As Long = OneMinute ^ 2
    Const OneDay As Long = OneHour * 24


b = Split(t, ", ")
If InStr(t, ", ") Then Time2sec = -1: Exit Function 'много времен, низзя суммировать

b = Split(t, ":")
Select Case UBound(b)
Case 3 'dd:hh:mm:ss
temp = Val(b(0)) * OneDay + Val(b(1)) * OneHour + Val(b(2)) * OneMinute + Val(b(3))
Case 2 'hh:mm:ss
temp = Val(b(0)) * OneHour + Val(b(1)) * OneMinute + Val(b(2))
Case 1 'mm:ss
temp = Val(b(0)) * OneMinute + Val(b(1))
Case 0 'ss
temp = Val(b(0))
End Select

'temp = temp * 100 '+ Str(txt_MS.Text)
Time2sec = temp

Exit Function

err:
Time2sec = -1
ToDebug "Err_T2S " & err.Description
End Function

Public Function FileName2Title(F As String) As String
'название фильма из названия файла
'c:\fff\film_super.avi => Film super
'если vts_ или video_ts - оставить полный

Dim tmp As String

If InStr(1, F, "video_ts", vbTextCompare) Then
    FileName2Title = F
    Exit Function
ElseIf InStr(1, F, "vts_", vbTextCompare) Then
    FileName2Title = F
    Exit Function
End If

tmp = GetNameFromPathAndName(F)
GetExtensionFromFileName tmp, tmp

If Len(F) < 1 Then FileName2Title = vbNullString: Exit Function

If InStr(tmp, "_") Then tmp = Replace(tmp, "_", " ")
'If InStr(tmp, ".") Then tmp = Replace(tmp, ".", " ")

Do While InStr(1, tmp, "  ")
    tmp = Replace(tmp, "  ", " ")
Loop
tmp = Trim(tmp)

'первая заглавная
If Len(tmp) > 1 Then
    tmp = (UCase$(left$(tmp, 1)) & LCase$(right$(tmp, Len(tmp) - 1)))
Else
    tmp = UCase$(tmp)
End If

FileName2Title = tmp
End Function


''''''''''''''''''''''''''''''''''''REM DUPS from ARRAY
Public Sub remdups(ByRef arr() As String)

'-----------------------------
' Coded by Olof Larsson
'-----------------------------

Dim c As Long, coun As Long
coun = UBound(arr)
Dim g As Long
c = 1
For g = 0 To coun
    If g + c > coun Then
        Exit For
    Else
        If arr(g) <> vbNullString Then
            If arr(g) = arr(g + c) Then
                arr(g + c) = vbNullString
                c = c + 1
                g = g - 1
            Else
                c = 1
            End If
        End If
    End If
Next g
c = 0
For g = 0 To coun
    If arr(g) = vbNullString Then
    Else
        arr(c) = arr(g)
        c = c + 1
    End If
Next g

On Error Resume Next
If c > 0 Then ReDim Preserve arr(c - 1)
End Sub

Private Sub TriQuickSortString2(ByRef sArray() As String, ByVal iSplit As Long, ByVal iMin As Long, ByVal iMax As Long)
Dim i As Long
Dim j As Long
Dim sTemp As String

' *NOTE* no checks are made in this function because it is used internally.
' Validity checks are made in the public function that calls this one.

If (iMax - iMin) > iSplit Then
    i = (iMax + iMin) / 2

    If sArray(iMin) > sArray(i) Then SwapStrings sArray(iMin), sArray(i)
    If sArray(iMin) > sArray(iMax) Then SwapStrings sArray(iMin), sArray(iMax)
    If sArray(i) > sArray(iMax) Then SwapStrings sArray(i), sArray(iMax)

    j = iMax - 1
    SwapStrings sArray(i), sArray(j)
    i = iMin
    CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(sArray(j)), 4   ' sTemp = sArray(j)

    Do
        Do
            i = i + 1
        Loop While sArray(i) < sTemp

        Do
            j = j - 1
        Loop While sArray(j) > sTemp

        If j < i Then Exit Do
        SwapStrings sArray(i), sArray(j)
    Loop

    SwapStrings sArray(i), sArray(iMax - 1)

    TriQuickSortString2 sArray, iSplit, iMin, j
    TriQuickSortString2 sArray, iSplit, i + 1, iMax
End If

' clear temp var (sTemp)
i = 0
CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(i), 4
End Sub

Public Sub TriQuickSortString(ByRef sArray() As String, Optional ByVal SortOrder As SortOrder = SortAscending)
Dim iLBound As Long
Dim iUBound As Long

iLBound = LBound(sArray)
iUBound = UBound(sArray)

' *NOTE*  the value 4 is VERY important here !!!
' DO NOT CHANGE 4 FOR A LOWER VALUE !!!
TriQuickSortString2 sArray, 4, iLBound, iUBound
InsertionSortString sArray, iLBound, iUBound

If SortOrder = SortDescending Then ReverseStringArray sArray
End Sub

Private Sub InsertionSortString(ByRef sArray() As String, ByVal iMin As Long, ByVal iMax As Long)
Dim i As Long
Dim j As Long
Dim sTemp As String

' *NOTE* no checks are made in this function because it is used internally.
' Validity checks are made in the public function that calls this one.

For i = iMin + 1 To iMax
    CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(sArray(i)), 4   ' sTemp = sArray(i)
    j = i

    Do While j > iMin
        If sArray(j - 1) <= sTemp Then Exit Do

        CopyMemory ByVal VarPtr(sArray(j)), ByVal VarPtr(sArray(j - 1)), 4  ' sArray(j) = sArray(j - 1)
        j = j - 1
    Loop

    CopyMemory ByVal VarPtr(sArray(j)), ByVal VarPtr(sTemp), 4      ' sArray(j) = sTemp
Next i

' clear temp var (sTemp)
i = 0
CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(i), 4
End Sub

Public Sub ReverseStringArray(ByRef sArray() As String)
Dim iLBound As Long
Dim iUBound As Long

iLBound = LBound(sArray)
iUBound = UBound(sArray)

While iLBound < iUBound
    SwapStrings sArray(iLBound), sArray(iUBound)

    iLBound = iLBound + 1
    iUBound = iUBound - 1
Wend
End Sub

Private Sub SwapStrings(ByRef s1 As String, ByRef s2 As String)
Dim i As Long

' StrPtr() returns 0 (NULL) if string is not initialized
' But StrPtr() is 5% faster than using CopyMemory, so I used that workaround, which is safe and fast.
i = StrPtr(s1)
If i = 0 Then CopyMemory ByVal VarPtr(i), ByVal VarPtr(s1), 4

CopyMemory ByVal VarPtr(s1), ByVal VarPtr(s2), 4
CopyMemory ByVal VarPtr(s2), i, 4
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''END REM DUPS from ARRAY

Public Sub SetFrmShowPicPicture(w As Integer)
'hScroll
'положить картинку на форму, если надо искуственно растянуть
'FormShowPic

Dim ResizeFlag1 As Boolean    'увеличивать ли искусственно ширину
Dim ResizeFlag2 As Boolean
Dim koef As Single
Dim FictWidth As Long

With FrmMain
    If w = 0 Then


        If (.PicTempHid(w).Width = 5280) And (.PicTempHid(w).Height = 8640) Then ResizeFlag1 = True    ' 352x576 > 768x576
        If (.PicTempHid(w).Width = 5280) And (.PicTempHid(w).Height = 7200) Then ResizeFlag2 = True    ' 352x480 > 640x480
        If (.PicTempHid(w).Width = 7200) And (.PicTempHid(w).Height = 7200) Then ResizeFlag2 = True    ' 480x480 > 640x480
        If (.PicTempHid(w).Width = 7200) And (.PicTempHid(w).Height = 8640) Then ResizeFlag1 = True    ' 480x576 > 768x576

        If Not addflag Then
            koef = GetAspectFromVideoString

            Select Case koef
            Case 1.333    '4:3
                If ResizeFlag1 Then    'pal
                    FictWidth = 768    '11520
                ElseIf ResizeFlag2 Then    'ntsc
                    FictWidth = 640
                End If
            Case 1.777    '16:9
                If ResizeFlag1 Then    'pal
                    FictWidth = 1024
                ElseIf ResizeFlag2 Then    'ntsc
                    FictWidth = 853
                End If
            End Select

        End If

        'On Error GoTo 0
        If Opt_CenterShowPic Then
            'центровать еще раз, для скролла (ShowInShowPic)
            frmShow_xPos = (Screen.Width - (.PicTempHid(w).Width)) \ 2    '+ FrmMain.Left
            frmShow_yPos = (Screen.Height - (.PicTempHid(w).Height)) \ 2    '+ FrmMain.Top
            If frmShow_xPos < 0 Then frmShow_xPos = 0
            If frmShow_yPos < 0 Then frmShow_yPos = 0
        End If

        If ResizeFlag1 Or ResizeFlag2 Then
            DoEvents    'неплохо
            FormShowPic.Move frmShow_xPos, frmShow_yPos, FictWidth * Screen.TwipsPerPixelX, .PicTempHid(w).Height
            ' FormShowPic.AutoRedraw = True
            ResizeWIA FormShowPic, FictWidth, .PicTempHid(w).Height / Screen.TwipsPerPixelY, .PicTempHid(w)
            ' FormShowPic.AutoRedraw = False
            ToDebug .PicTempHid(w).Width / Screen.TwipsPerPixelX & "x" & .PicTempHid(w).Height / Screen.TwipsPerPixelY & " > " & FictWidth & "x" & .PicTempHid(w).Height / Screen.TwipsPerPixelY
        Else
            FormShowPic.Move frmShow_xPos, frmShow_yPos, .PicTempHid(w).Width, .PicTempHid(w).Height
            FormShowPic.Picture = .PicTempHid(w).Image
            'Debug.Print "W:" & FormShowPic.Width & " H:" & FormShowPic.Height
        End If

    Else
        FormShowPic.Move frmShow_xPos, frmShow_yPos, .PicTempHid(w).Width, .PicTempHid(w).Height
        FormShowPic.Picture = .PicTempHid(w).Image
    End If

    FormShowPic.CurrentX = 3: FormShowPic.CurrentY = 3

End With
End Sub

Public Sub MakeNormal(hwnd As Long)
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Sub MakeTopMost(hwnd As Long)
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Function GetAspectFromVideoString() As Single
'1.333
'добываем из поля lv

Dim s As String
s = FrmMain.ListView.SelectedItem.SubItems(dbVideoInd)
GetAspectFromVideoString = 1.333

If InStr(s, "4:3") Then GetAspectFromVideoString = 1.333
If InStr(s, "4/3") Then GetAspectFromVideoString = 1.333

If InStr(s, "16:9") Then GetAspectFromVideoString = 1.777
If InStr(s, "16/9") Then GetAspectFromVideoString = 1.777

End Function
Public Function GetAspectFromTextVideo() As Currency    'Single
'1.333
'добываем из поля редактора TextVideoHid
'типа из ifo

'Dim s As String
GetAspectFromTextVideo = MMI_Format    'старый из MMInfo
If LCase$(right$(mpgName, 6)) = "_0.vob" Then Exit Function

With frmEditor
    If InStr(.TextVideoHid, "4:3") Then GetAspectFromTextVideo = 1.333
    If InStr(.TextVideoHid, "4/3") Then GetAspectFromTextVideo = 1.333

    If InStr(.TextVideoHid, "16:9") Then GetAspectFromTextVideo = 1.777
    If InStr(.TextVideoHid, "16/9") Then GetAspectFromTextVideo = 1.777
End With
End Function


Public Sub ResizeWIA(p As Object, new_w As Long, new_h As Long, Optional pIn As PictureBox, Optional aratio As Boolean, Optional rot As Integer)
Dim img As ImageFile
Dim IP As ImageProcess
Dim vec As Vector

On Error GoTo wiaerr

Set vec = New Vector
'Set img = New ImageFile
Set IP = New ImageProcess

If pIn Is Nothing Then
vec.BinaryData = Picture2Array(p)
Else
vec.BinaryData = Picture2Array(pIn)
End If

Set img = vec.ImageFile
'ресайз
IP.Filters.Add IP.FilterInfos("Scale").FilterID
IP.Filters(1).Properties("MaximumWidth").Value = new_w '640
IP.Filters(1).Properties("MaximumHeight").Value = new_h '480
IP.Filters(1).Properties("PreserveAspectRatio").Value = aratio

If rot <> 0 Then 'вращать
IP.Filters.Add IP.FilterInfos("RotateFlip").FilterID
IP.Filters(2).Properties("RotationAngle") = rot
End If

Set img = IP.Apply(img)

If Not img Is Nothing Then
     Set p.Picture = img.ARGBData.Picture(img.Width, img.Height)
End If

Exit Sub

wiaerr:
ToDebug "Err_WIA: " & err.Description
End Sub

Public Function FileExists(sfile As String) As Boolean
Dim tFnd As WIN32_FIND_DATA
If Len(sfile) = 0 Then Exit Function
FileExists = (FindFirstFile(sfile, tFnd) <> -1)
End Function

Public Function Str2Val(s As String) As Double
On Error Resume Next
Str2Val = CDbl(s)
If err Then
    Str2Val = Val(s)
    err.Clear
End If
End Function

Public Sub FillUserCombo(cbName As String, cbObj As Object)
'заполнить комбо из user.lng
Dim IniKeyCount As Integer
Dim IniKeysArr() As String
Dim temp As String
Dim i As Integer

'проверять в вызывающем коде If FileExists(userFile) Then
cbObj.Clear
If frmOptFlag Then FrmOptions.lstComboNames.Clear

IniKeyCount = GetKeyNames(cbName, userFile, IniKeysArr)
If IniKeyCount > -1 Then
    For i = 0 To IniKeyCount
        temp = VBGetPrivateProfileString(cbName, IniKeysArr(i), userFile)
        If Len(temp) <> 0 Then
            cbObj.AddItem temp 'значение
            If frmOptFlag Then FrmOptions.lstComboNames.AddItem IniKeysArr(i)  'и имя ключа
        End If
    Next i
End If
End Sub

Public Function GetAutoIniKeyName(lst As ListBox) As String
Dim i As Integer, j As Integer
Dim arr() As Long

'массив - список
'придумать название очередного ключа в секции
'найти максимальное целое из уже существующих имен в секции и вернуть, увеличив на 1
If lst.ListCount > 0 Then
    'переписать в массив все нумерик имена
    ReDim arr(0)
    For i = 0 To lst.ListCount - 1
        If IsNumeric(lst.List(i)) Then
            ReDim Preserve arr(j)
            arr(j) = Val(lst.List(i))
            j = j + 1
        End If
    Next i

GetAutoIniKeyName = CStr(ArrayMax(arr) + 1)

Else
    GetAutoIniKeyName = "1"
End If
End Function

Function ArrayMax(arr As Variant, Optional MaxIndex As Long) As Long
'поиск максимума в массиве
Dim i As Long, sFirst As Long, sLast As Long

sFirst = LBound(arr)
sLast = UBound(arr)

MaxIndex = sFirst
ArrayMax = arr(MaxIndex)

For i = sFirst + 1 To sLast
    If ArrayMax < arr(i) Then
        MaxIndex = i
        ArrayMax = arr(MaxIndex)
    End If
Next i
End Function

Public Function strMPEG() As String
'возвращает только "MPV2" или "MPV1" в зависимости от глобального MPGCodec

Select Case MPGCodec
Case "MPV2", "MPEG-2V"
    strMPEG = "MPV2"
Case "MPV1", "MPEG-1V"
    strMPEG = "MPV1"
Case Else
    strMPEG = vbNullString
    ToDebug "кодек: " & MPGCodec
End Select
End Function

Public Function sys_StrDel(str As String, StartPos As Long, DelCount As Long) As String
'удаление в str слева DelCount символов, начиная с StartPos включительно
Dim LStr, RStr As String
  On Error GoTo Error
  
  If Len(str) = 0 Then GoTo Error
  If StartPos > Len(str) Then GoTo Error
  If StartPos = 0 Then GoTo Error
  If StartPos <= 1 Then
    StartPos = 0
  Else
    StartPos = StartPos - 1
  End If
      
  If StartPos + DelCount > Len(str) Then
    DelCount = Len(str) - StartPos
  End If
  
  If StartPos > 0 Then
    LStr = Mid$(str, 1, StartPos)
  Else
    LStr = vbNullString
  End If
  
  RStr = Mid$(str, (StartPos + DelCount) + 1, Len(str) - (StartPos + DelCount))
  sys_StrDel = LStr & RStr
  
  Exit Function
  
Error:
'вернуть как было
'Debug.Print "err_sd: " & err.Description
sys_StrDel = str
End Function
Public Function sys_StrDelRev(str As String, StartPos As Long, DelCount As Long) As String
'реверс. передача в sys_StrDel и реверс
On Error GoTo Error

str = StrReverse(str)
str = sys_StrDel(str, StartPos, DelCount)
sys_StrDelRev = StrReverse(str)

Exit Function
  
Error:
'вернуть как было
'Debug.Print "err_sdr: " & err.Description
sys_StrDelRev = str
End Function

Public Function sys_InsNumsIncr(sIns As String, nIncr As Long) As String
'создать строку, с увеличившимся на nIncr числом, находящимся в строке sIns
'sIns вся строка с числом = lTxt & insNum & rTxt

Dim i As Integer
Dim insNum As Long    'число внутри строк lTxt и rTxt, проверить на целое
Dim nTxt As String    'строка insNum
Dim lTxt As String, rTxt As String    'левые и правые части (по бокам числа)
Dim lenIns As Long    'длина sIns
Dim FirstDigitFlag As Boolean    'флаг начала числа в строке
Dim StartDigit As Long, LenDigit As Long
Dim sShablon As String
Dim ch As String

On Error GoTo err

'разделить sTmp на lTxt & lNum & rTxt
lenIns = Len(sIns)
For i = 1 To lenIns
    ch = Mid$(sIns, i, 1)
    If IsNumeric(ch) Then  'с i началось число
        If Not FirstDigitFlag Then
            StartDigit = i
            LenDigit = lenIns - i + 1    'до конца строки
            FirstDigitFlag = True    'добыли начало числа
        End If
    Else
        If FirstDigitFlag Then    'если встречалось число
            LenDigit = i - StartDigit
            Exit For    'число кончилось, выйти
        End If
    End If
Next i
If StartDigit = 0 Then Exit Function 'нет числа, выход

nTxt = Mid$(sIns, StartDigit, LenDigit)    'nTxt - уже целое (070)
insNum = Val(nTxt) + nIncr 'увеличить на
'применить шаблоны
sShablon = String$(Len(nTxt), "0")
nTxt = Format$(insNum, sShablon)
lTxt = left$(sIns, StartDigit - 1)
rTxt = right$(sIns, lenIns - (Len(lTxt) + LenDigit))    '1
'убить 0 слева
lTxt = Replace(lTxt, "0", vbNullString)    '2

sys_InsNumsIncr = lTxt & nTxt & rTxt

Exit Function

err:
'не вставлять ничего
'Debug.Print "err_inn" & err.Description
sys_InsNumsIncr = vbNullString
End Function

Public Function UcaseCharAfterDelimiter(Expression As String, Delimiters As String) As String
'в верхний регистр все буквы после разделителя и первая, остальные - в нижний
    Dim cExp As Long, ubExpr As Long
    Dim cDel As Integer, ubDelim As Integer
    Dim aExpr() As Integer, aDelim() As Integer
    Dim sa1 As SAFEARRAY1D, sa2 As SAFEARRAY1D
 
    ubExpr = Len(Expression)
    ubDelim = Len(Delimiters)
    Expression = UCase$(left$(Expression, 1)) & LCase$(right$(Expression, ubExpr - 1)) 'вверх первый символ
    
    sa1.cbElements = 2:     sa1.cElements = ubExpr
    sa1.cDims = 1:          sa1.pvData = StrPtr(Expression)
    RtlMoveMemory ByVal VarPtrArray(aExpr), VarPtr(sa1), 4
    
    sa2.cbElements = 2:     sa2.cElements = ubDelim
    sa2.cDims = 1:          sa2.pvData = StrPtr(Delimiters)
    RtlMoveMemory ByVal VarPtrArray(aDelim), VarPtr(sa2), 4
    
    ubDelim = ubDelim - 1
    For cExp = 0 To ubExpr - 2
        For cDel = 0 To ubDelim
            If aExpr(cExp) = aDelim(cDel) Then
                    'On Error Resume Next
                    Mid$(Expression, cExp + 2, 1) = UCase$(Mid$(Expression, cExp + 2, 1))
            End If
        Next cDel
    Next cExp
  
    UcaseCharAfterDelimiter = Expression
    
    'чистка
    RtlZeroMemory ByVal VarPtrArray(aExpr), 4
    RtlZeroMemory ByVal VarPtrArray(aDelim), 4
End Function

Public Function CheckNoNull(F As String) As String
'принимает номер поля или имя
On Error Resume Next
If IsNumeric(F) Then
    If Not IsNull(rs.Fields(Val(F))) Then
        CheckNoNull = rs.Fields(Val(F))
    Else
        CheckNoNull = vbNullString
    End If
Else
    If Not IsNull(rs.Fields(F)) Then
        CheckNoNull = rs.Fields(F)
    Else
        CheckNoNull = vbNullString
End If
End If
End Function

Public Function CheckNoNullStr(F As String) As String
'принимает имя поля
On Error Resume Next
If Not IsNull(rs.Fields(F)) Then
        CheckNoNullStr = rs.Fields(F)
Else
        CheckNoNullStr = vbNullString
End If
End Function
Public Sub KillLdb(fn As String)
'прибить ldb файл
On Error Resume Next
If FileExists(fn) Then
    Kill fn
End If
End Sub
Public Function IsNotEmptyOrZero(s As String) As Boolean
'да, если есть текст и он не "0"

IsNotEmptyOrZero = True

If LenB(s) = 0 Then IsNotEmptyOrZero = False: Exit Function

If IsNumeric(s) Then
 If Val(s) = 0 Then IsNotEmptyOrZero = False: Exit Function
End If

End Function

' method 4 - the advanced **FAST** method для резервирования строки
'как быстрая замена Space$(lSize) String$(lSize, " ")
Public Function AllocString_ADV(ByVal lSize As Long) As String
RtlMoveMemory ByVal VarPtr(AllocString_ADV), _
SysAllocStringByteLen(0&, lSize + lSize), 4&
End Function



Public Sub SelectWordsGroup(tB As TextBox, st As Long)
Dim lPos As Long, rpos As Long, tmp As Long
Dim l_arr() As Long, r_arr() As Long
Dim s_arr As Variant
Dim i As Integer
Dim lenTXT As Long

On Error GoTo err

s_arr = Array("(", ")", "/", ",", "\", ";")
ReDim l_arr(UBound(s_arr))
ReDim r_arr(UBound(s_arr))
lenTXT = Len(tB.Text)

For i = 0 To UBound(s_arr)
    l_arr(i) = InStrRev(tB.Text, s_arr(i), st)
Next i
For i = 0 To UBound(s_arr)
    tmp = InStr(st, tB.Text, s_arr(i))
    If tmp > st Then r_arr(i) = tmp - 1 Else r_arr(i) = lenTXT
Next i

lPos = Max_In_Array(l_arr())    'от 0 до st
'поискать пробел справа от найденного и +1 если есть
Do While Mid$(tB.Text, lPos + 1, 1) = " "
    lPos = lPos + 1
Loop

rpos = Min_In_Array(r_arr())    'от st до конца строки
'поискать пробел слева от найденного и -1 если есть
Do While Mid$(tB.Text, rpos, 1) = " "
    rpos = rpos - 1
Loop

tB.SelStart = lPos
tB.SelLength = rpos - lPos

Exit Sub

err:
Debug.Print "Err_SWOGR"

End Sub

Private Function Max_In_Array(a() As Long) As Long
Dim i As Integer, Max As Long
For i = 0 To UBound(a)
If a(i) > Max Then Max = a(i)
Next i
Max_In_Array = Max
End Function
Private Function Min_In_Array(a() As Long) As Long
Dim i As Integer, min As Long
min = a(0)
For i = 1 To UBound(a)
If a(i) < min Then min = a(i)
Next i

Min_In_Array = min
End Function

Public Sub SaveHistory(sBaseName As String)
Dim tmp As String
'On Error Resume Next
    If FileExists(userFile) Then
        tmp = Join(arrHistoryKeys, ",")
        tmp = Replace(tmp, Kavs, vbNullString) 'убрали кавычку в конце каждого
        WriteKey "History", sBaseName, tmp, userFile
    End If
End Sub
Public Sub RestoreHistory(sBaseName As String)
Dim tmp As String
Dim R() As String
Dim i As Long, j As Long
Dim rsTmp As DAO.Recordset
Dim sSQL As String
Dim ClearFlag As Boolean

If rs Is Nothing Then Exit Sub

On Error Resume Next

'очистить
For i = 0 To nHistory
    arrHistoryKeys(i) = vbNullString
    arrHistoryTitles(i) = vbNullString
Next i

If FileExists(userFile) Then
    tmp = VBGetPrivateProfileString("History", sBaseName, userFile)
    If Len(tmp) <> 0 Then

        'tmp = Replace(tmp, Kavs, vbNullString) 'убрали кавычку в конце каждого
        If Tokenize04(tmp, R(), ",", False) > -1 Then
            tmp = Join(R, ",")
            sSQL = "Select MovieName,Key from Storage Where Key In (" & tmp & ")"
            Set rsTmp = DB.OpenRecordset(sSQL)

            'заполнить как кликали
            If Not (rsTmp.BOF And rsTmp.EOF) Then
                rsTmp.MoveLast
                For i = 0 To UBound(R)
                    arrHistoryKeys(i) = R(i) & Kavs
                    rsTmp.MoveFirst
                    For j = 0 To rsTmp.RecordCount - 1
                    'тут если в базе нет уже этой записи, то ключ будет, а названия нет
                    'т.е. будут пустышки, она будет не видна при сабклассе
                        If rsTmp("Key") = R(i) Then
                            arrHistoryTitles(i) = rsTmp("MovieName")
                            Exit For
                        End If
                        rsTmp.MoveNext
                    Next j
                Next i
            End If

            'так список будет как в базе, а не как кликали
            'хорошо: если после запоминания в историю фильм удалили, то тут найдется просто меньше фильмов
            '            If Not (rsTmp.BOF And rsTmp.EOF) Then
            '                rsTmp.MoveLast: rsTmp.MoveFirst
            '                For j = 0 To rsTmp.RecordCount - 1
            '                    arrHistoryTitles(j) = rsTmp("MovieName")
            '                    For i = 0 To UBound(R)
            '                        If rsTmp("Key") = R(i) Then
            '                            arrHistoryKeys(j) = R(i) & Kavs
            '                            Exit For
            '                        End If
            '                    Next i
            '                    rsTmp.MoveNext
            '                Next j
            '            End If



        End If
    End If
End If

Set rsTmp = Nothing
End Sub

'
'Private Function Simil(String1 As String, String2 As String) As Double   'Single
''req SubSim
'' приблтзительное равенство строк (1 = равны)
'Dim l1 As Long
'Dim l2 As Long
'String1 = UCase$(String1): String2 = UCase$(String2)
'    If String1 = String2 Then
'        Simil = 1
'    Else
'        l1 = Len(String1)
'        l2 = Len(String2)
'        If l1 <> 0 And l2 <> 0 Then
'            b1 = StrConv(String1, vbFromUnicode, LCID)
'            b2 = StrConv(String2, vbFromUnicode, LCID)
'            Simil = SubSim(1, l1, 1, l2) / (l1 + l2) * 2
'        End If
'    End If
'    'Erase b1 'потом в FuzzyDupsAct
'    'Erase b2
'End Function
'
'Private Function SubSim(st1 As Long, end1 As Long, st2 As Long, end2 As Long) As Long
''for Simil
'Dim c1 As Long
'Dim c2 As Long
'Dim ns1 As Long
'Dim ns2 As Long
'Dim i As Long
'Dim Max As Long
'
'If st1 > end1 Then Exit Function
'If st2 > end2 Then Exit Function
'If st1 <= 0 Then Exit Function
'If st2 <= 0 Then Exit Function
'
'For c1 = st1 To end1
'    For c2 = st2 To end2
'        i = 0
'        Do Until b1(c1 + i - 1) <> b2(c2 + i - 1)
'            i = i + 1
'            If i > Max Then
'                ns1 = c1
'                ns2 = c2
'                Max = i
'            End If
'            If c1 + i > end1 Or c2 + i > end2 Then Exit Do
'        Loop
'    Next
'Next
''If (end1 + end2) / 2 - Max < 5 Then 'продолжить, если достаточное совпадение букв
'Max = Max + SubSim(ns1 + Max, end1, ns2 + Max, end2)
'Max = Max + SubSim(st1, ns1 - 1, st2, ns2 - 1)
'SubSim = Max
''End If
'End Function
'Public Sub FuzzyDupsAct()
'
'Dim i As Integer
'Dim si As Double    'Single
'
'Dim arrKeys() As String    'с 0
'Dim Pers As String
'Dim sKeys As String
'Dim Itm As ListItem
'
'On Error GoTo err
'Screen.MousePointer = vbHourglass
'
''ReDim rsArr(0)
'
''ReDim arrKeys(FrmMain.LVActer.ListItems.Count)
''Dim j As Long, k As Long
''For j = 1 To FrmMain.LVActer.ListItems.Count
''    For k = j + 1 To FrmMain.LVActer.ListItems.Count
''    If Abs(Len(FrmMain.LVActer.ListItems(j)) - Len(FrmMain.LVActer.ListItems(k))) < 5 Then
''        si = Simil(FrmMain.LVActer.ListItems(j), FrmMain.LVActer.ListItems(k))
''        If si > 0.9 Then
''            'If si < 1 Then
''                'запомнить
''                i = i + 1
''                arrKeys(i) = Val(FrmMain.LVActer.ListItems(j).Key)
''                i = i + 1
''                arrKeys(i) = Val(FrmMain.LVActer.ListItems(k).Key)
''                Exit For
''            'End If
''        End If
''    End If
''    Next k
''Next j
'
'ReDim arrKeys(ars.RecordCount)
'ars.MoveFirst
'Do While Not ars.EOF
'DoEvents
'    If IsNull(ars("Name")) Then    'пустышка раз
'    ElseIf Len(ars("Name")) = 0 Then    'пустышка два
'    Else
'        i = i + 1
'        'и по всему списку
'        For Each Itm In FrmMain.LVActer.ListItems
'            If Abs(Len(ars("Name")) - Len(Itm.Text)) < 5 Then
'                si = Simil(ars("Name"), Itm.Text)
'                If si > 0.9 Then
'                    If si < 1 Then 'а то найдет самого себя, то есть будут все
'                        'запомнить
'                        arrKeys(i) = ars("Key")
'                        'FrmMain.LVActer.ListItems.Remove Itm.Index
'                        Exit For
'                    End If
'                End If
'            End If
'        Next
'    End If
'    ars.MoveNext
'Loop
'
''дубли, не нужно при count
'TriQuickSortString arrKeys            'sorts your string array
'remdups arrKeys                       'removes dups
'
'Screen.MousePointer = vbNormal
'If UBound(arrKeys) = 0 Then Exit Sub    'не нашли ничего
''ключи в строку
'sKeys = Join(arrKeys, ",")
''Debug.Print sKeys
'
''выдать из базы по ключам
'Set ars = ADB.OpenRecordset("Select * From Acter Where Key In (" & sKeys & ")")
''применить к списку
'ArsProcess
'Erase b1    'чистка глобальных
'Erase b2
'Exit Sub
'err:
'Screen.MousePointer = vbNormal
'Debug.Print err.Description
''ToDebug "Err_ActFuzz"
''MsgBox msgsvc(46), vbExclamation    ': ToDebug err.Description
'End Sub



'Public Function GetAspectRatio(ars As String, hs As String) As Single
'Dim temp As String
'temp = ars & hs
''16x9 pal = 1.459
''16x9 ntsc = 1.215
''4x3 pal = 1.094
''4x3 ntsc = 0.911
'Select Case temp
'Case "1576"
'GetAspectRatio = 1.094
'Case "1480"
'GetAspectRatio = 1.333 '1.215
'
'Case "4/3576"
'GetAspectRatio = 1.094
'Case "4/3288"
'GetAspectRatio = 1
'Case "4/3480"
'GetAspectRatio = 1.094 '1.215
'Case "4/3240"
'GetAspectRatio = 1 '0.911
'
'Case "16/9576"
'GetAspectRatio = 1.459
'Case "16/9288"
'GetAspectRatio = 1 '1.459
'Case "16/9480"
'GetAspectRatio = 1.215
'Case "16/9240"
'GetAspectRatio = 1 '1.215
'Case Else
'GetAspectRatio = 1 '1.333
'End Select
'
'End Function

'Private Function CheckANoNull(F As String) As String
''Dim temp As String
'If ars.Fields(F) <> vbNullString Then
'CheckANoNull = ars.Fields(F)
'Else
'CheckANoNull = vbNullString
'End If
'End Function
