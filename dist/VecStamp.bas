Attribute VB_Name = "VecStamp"
' =====================================================================
'  VecStamp - export selected PowerPoint shapes as vector graphics
'  Chinese name: Shi Yin ("vector stamp"). Project page and docs: README.md
'
'  Copyright (c) 2026 vluckyzhang <vluckyzhang@gmail.com>
'  Released under the MIT License. See the LICENSE file for details.
'
'  Notes for developers
'  - This is the readable UTF-8 source. build/build_bas.py turns it into
'    dist/VecStamp.bas: every Chinese string literal becomes U("hex") so the
'    module imports correctly under any Windows locale, and the generated
'    block at the end is filled with the VSDX template parts from src/vsdx
'    and the XPS -> SVG converter from src/helpers/xps2svg.ps1.
'  - All PowerPoint objects are late bound (As Object) and enum values are
'    numeric constants, so the module only depends on the VBA library and
'    compiles the same way on Windows and macOS.
'  - Ribbon callbacks are prefixed VS_ (see src/customUI14.xml).
'  - macOS: dialogs, moving files out of the PowerPoint sandbox, gzip, zip, the
'    colour picker, the progress dialog and the clipboard are done by the
'    AppleScript helper VecStamp.scpt that
'    the macOS installer puts in ~/Library/Application Scripts/com.microsoft.Powerpoint/.
' =====================================================================
Option Explicit

#If Mac Then
' (no Windows API declarations on macOS)
#Else
#If VBA7 Then
Private Type OPENFILENAMEW
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As LongPtr
    lpstrCustomFilter As LongPtr
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As LongPtr
    nMaxFile As Long
    lpstrFileTitle As LongPtr
    nMaxFileTitle As Long
    lpstrInitialDir As LongPtr
    lpstrTitle As LongPtr
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As LongPtr
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As LongPtr
    pvReserved As LongPtr
    dwReserved As Long
    FlagsEx As Long
End Type
Private Type CHOOSECOLORW
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    rgbResult As Long
    lpCustColors As LongPtr
    Flags As Long
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As LongPtr
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type INITCOMMONCONTROLSEX_T
    dwSize As Long
    dwICC As Long
End Type
Private Declare PtrSafe Function GetSaveFileNameW Lib "comdlg32.dll" (ByRef ofn As OPENFILENAMEW) As Long
Private Declare PtrSafe Function ChooseColorW Lib "comdlg32.dll" (ByRef cc As CHOOSECOLORW) As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, _
    ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, _
    ByVal lpParam As LongPtr) As LongPtr
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function SetWindowTextW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Function SendMessageW Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, _
    ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function CreateFontW Lib "gdi32" (ByVal nHeight As Long, ByVal nWidth As Long, _
    ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, _
    ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, _
    ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function InitCommonControlsEx Lib "comctl32" (ByRef icc As INITCOMMONCONTROLSEX_T) As Long
Private Type STARTUPINFOW
    cb As Long
    lpReserved As LongPtr
    lpDesktop As LongPtr
    lpTitle As LongPtr
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As LongPtr
    hStdInput As LongPtr
    hStdOutput As LongPtr
    hStdError As LongPtr
End Type
Private Type PROCESS_INFORMATION
    hProcess As LongPtr
    hThread As LongPtr
    dwProcessId As Long
    dwThreadId As Long
End Type
Private Declare PtrSafe Function CreateProcessW Lib "kernel32" (ByVal lpApplicationName As LongPtr, ByVal lpCommandLine As LongPtr, _
    ByVal lpProcessAttributes As LongPtr, ByVal lpThreadAttributes As LongPtr, ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, ByVal lpEnvironment As LongPtr, ByVal lpCurrentDirectory As LongPtr, _
    ByRef lpStartupInfo As STARTUPINFOW, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As LongPtr, ByRef lpExitCode As Long) As Long
Private Declare PtrSafe Function TerminateProcess Lib "kernel32" (ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
' colour panel (a native window driven by a message loop in ColorPanel)
Private Type MSG_T
    hwnd As LongPtr
    message As Long
    wParam As LongPtr
    lParam As LongPtr
    msgTime As Long
    ptX As Long
    ptY As Long
End Type
Private Declare PtrSafe Function PeekMessageW Lib "user32" (ByRef lpMsg As MSG_T, ByVal hwnd As LongPtr, _
    ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare PtrSafe Function TranslateMessage Lib "user32" (ByRef lpMsg As MSG_T) As Long
Private Declare PtrSafe Function DispatchMessageW Lib "user32" (ByRef lpMsg As MSG_T) As LongPtr
Private Declare PtrSafe Function IsDialogMessageW Lib "user32" (ByVal hDlg As LongPtr, ByRef lpMsg As MSG_T) As Long
Private Declare PtrSafe Function EnableWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal bEnable As Long) As Long
Private Declare PtrSafe Function SetFocusW Lib "user32" Alias "SetFocus" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function GetWindowTextW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpString As LongPtr, _
    ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function LoadImageW Lib "user32" (ByVal hInst As LongPtr, ByVal lpszName As LongPtr, _
    ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As LongPtr
Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDoubleClickTime Lib "user32" () As Long
Private Declare PtrSafe Function PostQuitMessage Lib "user32" (ByVal nExitCode As Long) As Long
Private Declare PtrSafe Sub SleepX Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function AdjustWindowRectEx Lib "user32" (ByRef lpRect As RECT, ByVal dwStyle As Long, _
    ByVal bMenu As Long, ByVal dwExStyle As Long) As Long
Private gPgWnd As LongPtr, gPgMsg As LongPtr, gPgBar As LongPtr, gPgSub As LongPtr, gPgFont As LongPtr, gPgFontB As LongPtr
Private gPgPct As LongPtr
#Else
Private Type OPENFILENAMEW
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As Long
    lpstrCustomFilter As Long
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As Long
    nMaxFile As Long
    lpstrFileTitle As Long
    nMaxFileTitle As Long
    lpstrInitialDir As Long
    lpstrTitle As Long
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
    pvReserved As Long
    dwReserved As Long
    FlagsEx As Long
End Type
Private Type CHOOSECOLORW
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type INITCOMMONCONTROLSEX_T
    dwSize As Long
    dwICC As Long
End Type
Private Declare Function GetSaveFileNameW Lib "comdlg32.dll" (ByRef ofn As OPENFILENAMEW) As Long
Private Declare Function ChooseColorW Lib "comdlg32.dll" (ByRef cc As CHOOSECOLORW) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, _
    ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, _
    ByVal lpParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowTextW Lib "user32" (ByVal hwnd As Long, ByVal lpString As Long) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateFontW Lib "gdi32" (ByVal nHeight As Long, ByVal nWidth As Long, _
    ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, _
    ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, _
    ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32" (ByRef icc As INITCOMMONCONTROLSEX_T) As Long
Private Type STARTUPINFOW
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
Private Declare Function CreateProcessW Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, _
    ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
    ByRef lpStartupInfo As STARTUPINFOW, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Type MSG_T
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    msgTime As Long
    ptX As Long
    ptY As Long
End Type
Private Declare Function PeekMessageW Lib "user32" (ByRef lpMsg As MSG_T, ByVal hwnd As Long, _
    ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (ByRef lpMsg As MSG_T) As Long
Private Declare Function DispatchMessageW Lib "user32" (ByRef lpMsg As MSG_T) As Long
Private Declare Function IsDialogMessageW Lib "user32" (ByVal hDlg As Long, ByRef lpMsg As MSG_T) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal bEnable As Long) As Long
Private Declare Function SetFocusW Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetWindowTextW Lib "user32" (ByVal hwnd As Long, ByVal lpString As Long, _
    ByVal nMaxCount As Long) As Long
Private Declare Function LoadImageW Lib "user32" (ByVal hInst As Long, ByVal lpszName As Long, _
    ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetDoubleClickTime Lib "user32" () As Long
Private Declare Function PostQuitMessage Lib "user32" (ByVal nExitCode As Long) As Long
Private Declare Sub SleepX Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Private Declare Function AdjustWindowRectEx Lib "user32" (ByRef lpRect As RECT, ByVal dwStyle As Long, _
    ByVal bMenu As Long, ByVal dwExStyle As Long) As Long
Private gPgWnd As Long, gPgMsg As Long, gPgBar As Long, gPgSub As Long, gPgFont As Long, gPgFontB As Long
Private gPgPct As Long
#End If
#End If

' ---- product information
Private Const VS_NAME As String = "VecStamp"
Private Const VS_VERSION As String = "1.2.2"
Private Const VS_AUTHOR As String = "vluckyzhang"
Private Const VS_EMAIL As String = "vluckyzhang@gmail.com"
Private Const VS_YEAR As String = "2026"
Private Const VS_REPO As String = "https://github.com/vluckyzhang/VecStamp"
Private Const REG_ROOT As String = "HKCU\Software\VecStamp\"
Private Const MAC_SCRIPT As String = "VecStamp.scpt"
Private Const MAC_SEP As String = "|*|"

' ---- PowerPoint / Office enum values (numeric: late binding, no type-library dependency)
Private Const PP_FMT_PNG As Long = 2
Private Const PP_FMT_WMF As Long = 4
Private Const PP_FMT_EMF As Long = 5
Private Const PP_FMT_SVG As Long = 6
Private Const PP_RELATIVE_TO_SLIDE As Long = 1
Private Const PP_SEL_SLIDES As Long = 1
Private Const PP_SEL_SHAPES As Long = 2
Private Const PP_SEL_TEXT As Long = 3
Private Const PP_PASTE_EMF As Long = 2
Private Const PP_LAYOUT_BLANK As Long = 12
Private Const PP_SAVE_AS_PDF As Long = 32
Private Const MSO_TRUE As Long = -1
Private Const MSO_FALSE As Long = 0
Private Const MSO_RECTANGLE As Long = 1
Private Const MSO_AUTOSHAPE As Long = 1
Private Const MSO_CALLOUT As Long = 2
Private Const MSO_FREEFORM As Long = 5
Private Const MSO_GROUP As Long = 6
Private Const MSO_LINE As Long = 9
Private Const MSO_TEXTEFFECT As Long = 15
Private Const MSO_TEXTBOX As Long = 17
Private Const KAPPA As Double = 0.552284749830794
Private Const PI2 As Double = 6.28318530717959
Private Const PDEG As Double = 2.90888208665722E-07     ' radians per DrawingML angle unit (1/60000 degree)
Private Const MSO_SEND_TO_BACK As Long = 1
Private Const MSO_PLACEHOLDER As Long = 14
Private Const MSO_FOLDER_PICKER As Long = 4

' ---- our format index (= order in the ribbon drop-down and the save dialog)
Private Const FMT_EMF As Long = 0
Private Const FMT_EMZ As Long = 1
Private Const FMT_SVG As Long = 2
Private Const FMT_WMF As Long = 3
Private Const FMT_PNG As Long = 4
Private Const FMT_PDF As Long = 5
Private Const FMT_VSDX As Long = 6
Private Const FMT_LAST As Long = 6

' ---- backgrounds (= order in the ribbon drop-down)
Private Const BG_NONE As Long = 0
Private Const BG_WHITE As Long = 1
Private Const BG_BEIGE As Long = 2
Private Const BG_GRID As Long = 3
Private Const BG_CUSTOM As Long = 4
Private Const DEF_CUSTOM_COLOR As String = "255,248,235"

Private Const DPI_MAX As Long = 3000
Private Const SVG_TIMEOUT As Long = 180          ' seconds for the built-in XPS -> SVG conversion
Private Const PNG_WARN_PIXELS As Double = 120000000#

#If Mac Then
#Else
Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_HIDEREADONLY As Long = &H4
Private Const OFN_NOCHANGEDIR As Long = &H8
Private Const OFN_PATHMUSTEXIST As Long = &H800
Private Const OFN_EXPLORER As Long = &H80000
Private Const CC_RGBINIT As Long = &H1
Private Const CC_FULLOPEN As Long = &H2
Private Const CC_ANYCOLOR As Long = &H100
Private Const WS_POPUP As Long = &H80000000
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_EX_DLGMODALFRAME As Long = &H1
Private Const WS_EX_TOOLWINDOW As Long = &H80
Private Const SS_NOPREFIX As Long = &H80
Private Const WM_SETFONT As Long = &H30
Private Const PBM_SETPOS As Long = &H402
Private Const PBM_SETRANGE32 As Long = &H406
Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNOACTIVATE As Long = 4
Private Const LOGPIXELSY As Long = 90
Private Const VK_ESCAPE As Long = &H1B
Private Const ICC_PROGRESS_CLASS As Long = &H20
Private Const SS_RIGHT As Long = &H2
Private Const CREATE_NO_WINDOW As Long = &H8000000
Private Const STARTF_USESHOWWINDOW As Long = &H1
Private Const WAIT_TIMEOUT As Long = &H102
Private Const PM_REMOVE As Long = 1
Private Const WM_QUIT As Long = &H12
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const VK_RETURN As Long = &HD
Private Const VK_SPACE As Long = &H20
Private Const WS_TABSTOP As Long = &H10000
Private Const WS_BORDER As Long = &H800000
Private Const WS_EX_CONTROLPARENT As Long = &H10000
Private Const SS_NOTIFY As Long = &H100
Private Const SS_BITMAP As Long = &HE
Private Const STM_SETIMAGE As Long = &H172
Private Const IMAGE_BITMAP As Long = 0
Private Const LR_LOADFROMFILE As Long = &H10
Private Const BS_DEFPUSHBUTTON As Long = &H1
Private Const ES_AUTOHSCROLL As Long = &H80
Private Const EM_LIMITTEXT As Long = &HC5
Private Const EM_SETSEL As Long = &HB1
Private Const WS_EX_CLIENTEDGE As Long = &H200
Private Const SW_SHOW As Long = 5
Private Const COLOR_BTNFACE As Long = 15
#End If

' ---- results of RunHidden besides the program's own exit code
Private Const RUN_FAILED As Long = -1
Private Const RUN_CANCELLED As Long = -2
Private Const RUN_TIMEOUT As Long = -3
Private Const ERR_CANCEL As Long = vbObjectError + 599

Private gRibbon As Object
Private gLang As String             ' interface language (zh-CN, zh-TW, en, ja)
Private gStr As Collection          ' key -> text in gLang
Private gSvgFallback As Long        ' per run: 0 = not asked yet, 1 = save EMF instead, 2 = skip
Private gSvgNative As Long          ' per run: 0 = untested, 1 = PowerPoint writes valid SVG, 2 = it does not
Private gSvgXpsOff As Boolean       ' per run: the XPS -> SVG conversion timed out, use the other methods
Private gBigPngOk As Boolean        ' per run: user accepted a very large PNG
Private gCustColors(0 To 15) As Long
Private gBusy As Boolean            ' an export is running (ribbon clicks during DoEvents are ignored)
' progress window: items (shapes, files or slides) and the position inside the current item
Private gPgActive As Boolean, gPgItems As Long, gPgItem As Long, gPgFrac As Double, gPgTitle As String
Private gPgLabel As String, gPgDetail As String, gPgCancel As Boolean, gPgFile As String
#If Mac Then
Private gPgMacPid As String, gPgMacPct As Long, gPgMacTime As Double, gPgMacText As String, gPgMacCheck As Double
#End If
' native VSDX progress
Private gVxTotal As Long, gVxDone As Long
' native VSDX writer
Private gVxSb() As String, gVxSbN As Long, gVxId As Long, gVxMedia As Collection
Private gVxScratch As Object, gVxPres As Object, gVxNative As Long, gVxPictures As Long, gVxCopy As Long
Private gGeo As String, gGeoSec As String, gGeoIx As Long, gGeoRow As Long
Private gGeoW As Double, gGeoH As Double, gGeoFlipH As Boolean, gGeoFlipV As Boolean
Private gGeoX0 As Double, gGeoY0 As Double, gGeoLX As Double, gGeoLY As Double
Private gGeoNoFill As Boolean, gGeoNoLine As Boolean, gGeoOpenOnly As Boolean, gGeoTextBox As Boolean
Private gGeoCU As Double, gGeoCV As Double, gGeoSU As Double, gGeoSV As Double, gGeoClosed As Boolean
' preset geometry engine: guide values, adjust values (DrawingML units), shape size
Private gPv() As Double, gPa() As Double, gPw As Double, gPh As Double


' =====================================================================
'  Public macros (can be bound to the Quick Access Toolbar or run by name)
' =====================================================================
Public Sub VS_ExportDefault()
    DoExport GetIdx("Format", FMT_LAST)
End Sub

Public Sub VS_ExportEMF()
    DoExport FMT_EMF
End Sub

Public Sub VS_ExportEMZ()
    DoExport FMT_EMZ
End Sub

Public Sub VS_ExportSVG()
    DoExport FMT_SVG
End Sub

Public Sub VS_ExportWMF()
    DoExport FMT_WMF
End Sub

Public Sub VS_ExportPNG()
    DoExport FMT_PNG
End Sub

Public Sub VS_ExportPDF()
    DoExport FMT_PDF
End Sub

Public Sub VS_ExportVSDX()
    DoExport FMT_VSDX
End Sub

Public Sub VS_ExportSlides()
    DoBatch
End Sub

Public Sub VS_PictureCopy()
    DoPictureCopy
End Sub

Public Sub VS_PickColor()
    DoPickColor
End Sub

Public Sub VS_MoreColors()
    DoMoreColors
End Sub

Public Sub VS_Eyedropper()
    DoEyedropper
End Sub

Public Sub VS_Feedback()
    DoFeedback
End Sub


' =====================================================================
'  Ribbon callbacks
'  Every text comes from the string tables ("rb.<control id>" = label, ".tip" = screentip,
'  ".sup" = supertip, ".desc" = description), so the whole tab switches language at once when
'  the ribbon is invalidated (VS_OnCheck on a language entry).
' =====================================================================
Public Sub VS_OnLoad(ribbon As Object)
    Set gRibbon = ribbon
End Sub

Public Sub VS_OnExport(control As Object)
    Select Case control.Id
        Case "vsBtnEMF", "vsMnuEMF": DoExport FMT_EMF
        Case "vsBtnEMZ", "vsMnuEMZ": DoExport FMT_EMZ
        Case "vsBtnSVG", "vsMnuSVG": DoExport FMT_SVG
        Case "vsBtnWMF", "vsMnuWMF": DoExport FMT_WMF
        Case "vsBtnPNG", "vsMnuPNG": DoExport FMT_PNG
        Case "vsBtnPDF", "vsMnuPDF": DoExport FMT_PDF
        Case "vsBtnVSDX", "vsMnuVSDX": DoExport FMT_VSDX
        Case Else: DoExport GetIdx("Format", FMT_LAST)
    End Select
End Sub

Public Sub VS_OnBatch(control As Object)
    DoBatch
End Sub

Public Sub VS_OnPicture(control As Object)
    DoPictureCopy
End Sub

Public Sub VS_OnPickColor(control As Object)
    DoPickColor
End Sub

Public Sub VS_OnMoreColors(control As Object)
    DoMoreColors
End Sub

Public Sub VS_OnEyedropper(control As Object)
    DoEyedropper
End Sub

' ---- texts
Public Sub VS_GetLabel(control As Object, ByRef label)
    Dim id As String
    id = control.Id
    Select Case id
        Case "vsBtnExport"
            label = Tx("rb.vsBtnExport", UCase$(FmtExt(GetIdx("Format", FMT_LAST))))
        Case "vsMnuBg"
            label = Tx("rb.vsMnuBg", BgShortName(BgTypeIdx()))
        Case "vsGalColor", "vsBgCustomMac"
            label = Tx("rb." & id)
            If BgTypeIdx() = BG_CUSTOM Then label = label & "  " & ChrW$(&H2713)
        Case "vsMnuLang"
            label = LangName(CurLang())
        Case "vsMnuEMF", "vsMnuEMZ", "vsMnuSVG", "vsMnuWMF", "vsMnuPDF", "vsMnuVSDX", "vsMnuPNG"
            label = Tx("fmt." & LCase$(Mid$(id, 6)))
        Case Else
            label = Tx("rb." & id)
    End Select
End Sub

Public Sub VS_GetScreentip(control As Object, ByRef tip)
    Dim id As String
    id = control.Id
    If Left$(id, 6) = "vsBgSw" Then id = "vsMnuBg"
    If TxHas("rb." & id & ".tip") Then tip = Tx("rb." & id & ".tip") Else VS_GetLabel control, tip
End Sub

Public Sub VS_GetSupertip(control As Object, ByRef tip)
    Dim id As String
    id = control.Id
    If Left$(id, 6) = "vsBgSw" Then id = "vsMnuBg"
    Select Case id
        Case "vsBtnFolder": tip = Tx("rb.vsBtnFolder.sup", FolderSettingText())
        Case "vsBtnOpenFolder": tip = Tx("rb.vsBtnOpenFolder.sup", FileManagerName())
        Case "vsGalColor", "vsEbRgb", "vsBgCustomMac": tip = Tx("rb.vsGalColor.sup", ColorSettingText())
        Case "vsMnuBg": tip = Tx("rb.vsMnuBg.sup", BgName(BgTypeIdx()))
        Case "vsMnuLang": tip = Tx("rb.vsMnuLang.sup", LangName(CurLang()))
        Case Else
            If TxHas("rb." & id & ".sup") Then tip = Tx("rb." & id & ".sup") Else tip = SettingsSummary()
    End Select
End Sub

Public Sub VS_GetDescription(control As Object, ByRef desc)
    If TxHas("rb." & control.Id & ".desc") Then desc = Tx("rb." & control.Id & ".desc") Else desc = ""
End Sub

Public Sub VS_GetTitle(control As Object, ByRef title)
    title = Tx("rb." & control.Id)
End Sub

' ---- background: preview square (Windows: drawn; macOS: one of the static icons) and menu
Public Sub VS_GetVisible(control As Object, ByRef visible)
    Dim mac As Boolean, bg As Long
#If Mac Then
    mac = True
#End If
    bg = BgTypeIdx()
    Select Case control.Id
        Case "vsBgSwatch", "vsGalColor": visible = Not mac
        Case "vsBgCustomMac": visible = mac
        Case "vsBgSwNone": visible = mac And bg = BG_NONE
        Case "vsBgSwWhite": visible = mac And bg = BG_WHITE
        Case "vsBgSwBeige": visible = mac And bg = BG_BEIGE
        Case "vsBgSwGrid": visible = mac And bg = BG_GRID
        Case "vsBgSwCustom": visible = mac And bg = BG_CUSTOM
        Case Else: visible = True
    End Select
End Sub

Public Sub VS_GetImage(control As Object, ByRef image)
#If Mac Then
#Else
    Select Case control.Id
        Case "vsBgSwatch": Set image = BgPicture(BgTypeIdx(), 16)
        Case "vsGalColor": Set image = BgPicture(BG_CUSTOM, 32)
    End Select
#End If
End Sub

' ---- background colour gallery (theme / standard / recent colours), Windows
Public Sub VS_GalColorCount(control As Object, ByRef count)
    Dim rc As Variant
    rc = RecentColors()
    count = 70 + UBound(rc) + 1
End Sub

Public Sub VS_GalColorImage(control As Object, index As Integer, ByRef image)
#If Mac Then
#Else
    Set image = SwatchPicture(GalColor(index), 16)
#End If
End Sub

Public Sub VS_GalColorTip(control As Object, index As Integer, ByRef tip)
    Dim label As String
    GalColor index, label
    tip = label
End Sub

Public Sub VS_GalColorPick(control As Object, id As String, index As Integer)
    SetBgColor GalColor(index)
End Sub

' ---- drop-downs with translated items
Public Sub VS_GetItemCount(control As Object, ByRef count)
    Select Case control.Id
        Case "vsDdMode", "vsDdNaming": count = 2
        Case "vsDdDpi": count = 4
        Case Else: count = 0
    End Select
End Sub

Public Sub VS_GetItemLabel(control As Object, index As Integer, ByRef label)
    Select Case control.Id
        Case "vsDdMode": label = Tx(IIf(index = 0, "mode.merge", "mode.split"))
        Case "vsDdNaming": label = Tx(IIf(index = 0, "naming.ask", "naming.auto"))
        Case "vsDdDpi"
            If index < 3 Then label = CStr(Array(150, 300, 600)(index)) & " DPI" Else label = Tx("dpi.custom")
        Case Else: label = ""
    End Select
End Sub

Public Sub VS_GetItemID(control As Object, index As Integer, ByRef id)
    id = control.Id & "_" & index
End Sub

' ---- donations, project page
Public Sub VS_OnDonate(control As Object, id As String, index As Integer)
    If index = 0 Then DoDonateThanks Tx("don.alipay") Else DoDonateThanks Tx("don.wechat")
End Sub

Public Sub VS_OnDonateInfo(control As Object)
    DoAbout
End Sub

Public Sub VS_OnHome(control As Object)
    If Not OpenUrlX(VS_REPO) Then MsgBox Tx("msg.home", VS_REPO), vbInformation, AppTitle()
End Sub

Public Sub VS_OnOpenFolder(control As Object)
    Dim f As String
    f = GetSet("LastFolder", "")
    If Len(f) = 0 Or Not FolderExistsX(f) Then f = GetSet("Folder", "")
    If Len(f) = 0 Or Not FolderExistsX(f) Then
        MsgBox Tx("msg.noFolderYet"), vbInformation, AppTitle()
        Exit Sub
    End If
    OpenPathX f
End Sub

Public Sub VS_GetIndex(control As Object, ByRef index)
    Select Case control.Id
        Case "vsDdFormat": index = GetIdx("Format", FMT_LAST)
        Case "vsDdMode": index = GetIdx("Mode", 1)
        Case "vsDdNaming": index = GetIdx("Naming", 1)
        Case "vsDdDpi": index = GetIdx("Dpi", 3, 1)
        Case Else: index = 0
    End Select
End Sub

Public Sub VS_OnDropDown(control As Object, id As String, index As Integer)
    Select Case control.Id
        Case "vsDdFormat": PutSet "Format", CStr(index)
        Case "vsDdMode": PutSet "Mode", CStr(index)
        Case "vsDdNaming": PutSet "Naming", CStr(index)
        Case "vsDdDpi": PutSet "Dpi", CStr(index)
    End Select
    RefreshRibbon
End Sub

' ---- check boxes and toggle buttons (background entries, languages)
Public Sub VS_GetPressed(control As Object, ByRef returnedVal)
    Select Case control.Id
        Case "vsChkOpen": returnedVal = GetFlag("OpenFolder", "0")
        Case "vsChkVsdx": returnedVal = GetFlag("VsdxNative", "1")
        Case "vsBgNone": returnedVal = (BgTypeIdx() = BG_NONE)
        Case "vsBgWhite": returnedVal = (BgTypeIdx() = BG_WHITE)
        Case "vsBgBeige": returnedVal = (BgTypeIdx() = BG_BEIGE)
        Case "vsBgGrid": returnedVal = (BgTypeIdx() = BG_GRID)
        Case "vsBgCustomMac": returnedVal = (BgTypeIdx() = BG_CUSTOM)
        Case "vsLangZhCN": returnedVal = (CurLang() = "zh-CN")
        Case "vsLangZhTW": returnedVal = (CurLang() = "zh-TW")
        Case "vsLangEn": returnedVal = (CurLang() = "en")
        Case "vsLangJa": returnedVal = (CurLang() = "ja")
        Case Else: returnedVal = False
    End Select
End Sub

Public Sub VS_OnCheck(control As Object, pressed As Boolean)
    Dim v As String
    If pressed Then v = "1" Else v = "0"
    Select Case control.Id
        Case "vsChkOpen": PutSet "OpenFolder", v
        Case "vsChkVsdx": PutSet "VsdxNative", v
        Case "vsBgNone": PutSet "BgType", CStr(BG_NONE)
        Case "vsBgWhite": PutSet "BgType", CStr(BG_WHITE)
        Case "vsBgBeige": PutSet "BgType", CStr(BG_BEIGE)
        Case "vsBgGrid": PutSet "BgType", CStr(BG_GRID)
        Case "vsBgCustomMac": DoMoreColors
        Case "vsLangZhCN": SetLang "zh-CN"
        Case "vsLangZhTW": SetLang "zh-TW"
        Case "vsLangEn": SetLang "en"
        Case "vsLangJa": SetLang "ja"
    End Select
    RefreshRibbon
End Sub

Public Sub VS_GetEnabled(control As Object, ByRef enabled)
    Select Case control.Id
        Case "vsEbDpi": enabled = (GetIdx("Dpi", 3, 1) = 3)
        Case "vsEbGrid": enabled = (BgTypeIdx() = BG_GRID)
        Case "vsEbAlpha": enabled = (BgTypeIdx() <> BG_NONE)
        Case Else: enabled = True
    End Select
End Sub

Public Sub VS_GetText(control As Object, ByRef strText)
    Select Case control.Id
        Case "vsEbMargin": strText = NumText(GetMargin())
        Case "vsEbDpi": strText = CStr(CustomDpi())
        Case "vsEbRgb": strText = ColorSettingText()
        Case "vsEbAlpha": strText = CStr(BgAlpha())
        Case "vsEbGrid": strText = NumText(GridStep())
        Case Else: strText = ""
    End Select
End Sub

Public Sub VS_OnChange(control As Object, strText As String)
    Dim v As Double, r As Long, g As Long, b As Long
    Select Case control.Id
        Case "vsEbMargin"
            If ReadNumber(strText, 0, 500, v) Then PutSet "Margin", NumText(v) _
                Else MsgBox Tx("msg.badMargin"), vbExclamation, AppTitle()
        Case "vsEbDpi"
            If ReadNumber(strText, 1, DPI_MAX, v) Then
                PutSet "DpiCustom", CStr(CLng(v))
            Else
                MsgBox Tx("msg.badDpi", DPI_MAX), vbExclamation, AppTitle()
            End If
        Case "vsEbAlpha"
            If ReadNumber(strText, 0, 100, v) Then PutSet "BgAlpha", CStr(CLng(v)) _
                Else MsgBox Tx("msg.badAlpha"), vbExclamation, AppTitle()
        Case "vsEbGrid"
            If ReadNumber(strText, 2, 200, v) Then PutSet "GridStep", NumText(v) _
                Else MsgBox Tx("msg.badGrid"), vbExclamation, AppTitle()
        Case "vsEbRgb"
            If ParseColor(strText, r, g, b) Then
                SetBgColor RGB(r, g, b)
            Else
                MsgBox Tx("col.bad"), vbExclamation, AppTitle()
            End If
    End Select
    RefreshRibbon
End Sub

Public Sub VS_OnFolder(control As Object)
    Dim cur As String, f As String
    cur = GetSet("Folder", "")
    If Len(cur) = 0 Or Not FolderExistsX(cur) Then cur = DesktopPathX()
    f = PickFolderX(cur, Tx("set.pickDefaultFolder"))
    If Len(f) > 0 Then
        PutSet "Folder", f
        RefreshRibbon
    End If
End Sub

Public Sub VS_OnReset(control As Object)
    If MsgBox(Tx("msg.askReset"), vbQuestion + vbYesNo, AppTitle()) <> vbYes Then Exit Sub
    ResetSettings
    RefreshRibbon
End Sub

Public Sub VS_OnSettings(control As Object)
    MsgBox SettingsSummary(), vbInformation, AppTitle() & " - " & Tx("set.current")
End Sub

Public Sub VS_OnAbout(control As Object)
    DoAbout
End Sub

Public Sub VS_OnFeedback(control As Object)
    DoFeedback
End Sub

Public Sub VS_OnHelp(control As Object)
    MsgBox Tx("help.ppt", DPI_MAX), vbInformation, AppTitle()
End Sub


' =====================================================================
'  Main flows
' =====================================================================
Private Sub DoExport(ByVal fmt As Long)
    Dim win As Object, sr As Object, pres As Object
    Dim files As New Collection
    Dim path As String, folder As String, baseName As String, p As String
    Dim i As Long, autoSave As Boolean, separate As Boolean

    If gBusy Then Exit Sub
    On Error GoTo EH
    If Not GetTarget(win, sr, True) Then Exit Sub
    Set pres = win.Presentation
    ResetRunState
    autoSave = (GetIdx("Naming", 1) = 1)
    separate = (GetIdx("Mode", 1) = 1) And (sr.Count > 1)
    baseName = BaseNameFor(pres, sr)

    If separate Then
        If autoSave Then
            folder = DefaultFolder(pres)
        Else
            folder = PickFolderX(InitialFolder(pres), Tx("dlg.pickFolderEach"))
            If Len(folder) = 0 Then Exit Sub
        End If
        gBusy = True
        ProgressBegin Tx("prog.exportN", sr.Count, UCase$(FmtExt(fmt))), sr.Count
        For i = 1 To sr.Count
            ProgressItem i, Tx("prog.item", i, sr.Count, ShortName(sr.Item(i).Name))
            CheckCancel
            path = UniquePath(folder, baseName & "_" & SafeName(sr.Item(i).Name), FmtExt(fmt))
            p = ExportItem(pres, sr.Item(i), path, fmt)
            If Len(p) > 0 Then files.Add p
        Next i
    Else
        If autoSave Then
            path = UniquePath(DefaultFolder(pres), baseName & "_" & Format$(Now, "yyyymmdd_hhnnss"), FmtExt(fmt))
        Else
            path = AskSavePath(InitialFolder(pres), baseName, fmt)
            If Len(path) = 0 Then Exit Sub
        End If
        gBusy = True
        ProgressBegin Tx("prog.export1", UCase$(FmtExt(fmt))), 1
        ProgressItem 1, ""
        p = ExportItem(pres, sr, path, fmt)
        If Len(p) > 0 Then files.Add p
    End If

    ProgressEnd
    gBusy = False
    Report files, autoSave Or separate
    Exit Sub
EH:
    ProgressEnd
    gBusy = False
    If Err.Number = ERR_CANCEL Then
        ReportCancelled files
    Else
        MsgBox Tx("msg.exportFailed", Err.Description), vbExclamation, AppTitle()
    End If
End Sub

' One file per slide: all slides, or only the slides selected in the thumbnail pane.
Private Sub DoBatch()
    Dim win As Object, pres As Object, sld As Object, rng As Object
    Dim slideList As New Collection, files As New Collection
    Dim fmt As Long, i As Long, ans As VbMsgBoxResult
    Dim folder As String, baseName As String, path As String, p As String, skipped As Long

    On Error Resume Next
    Set win = Application.ActiveWindow
    On Error GoTo EH
    If win Is Nothing Then
        MsgBox Tx("msg.openPres"), vbInformation, AppTitle()
        Exit Sub
    End If
    Set pres = win.Presentation
    fmt = GetIdx("Format", FMT_LAST)

    If win.Selection.Type = PP_SEL_SLIDES Then
        If win.Selection.SlideRange.Count > 1 Then
            ans = MsgBox(Tx("msg.askWhich", win.Selection.SlideRange.Count, pres.Slides.Count), _
                         vbQuestion + vbYesNoCancel, AppTitle())
            If ans = vbCancel Then Exit Sub
            If ans = vbYes Then
                For i = 1 To win.Selection.SlideRange.Count
                    slideList.Add win.Selection.SlideRange.Item(i)
                Next i
            End If
        End If
    End If
    If slideList.Count = 0 Then
        For i = 1 To pres.Slides.Count
            slideList.Add pres.Slides.Item(i)
        Next i
    End If
    If slideList.Count = 0 Then
        MsgBox Tx("msg.noSlides"), vbInformation, AppTitle()
        Exit Sub
    End If

    If GetIdx("Naming", 1) = 1 Then
        folder = DefaultFolder(pres)
    Else
        folder = PickFolderX(InitialFolder(pres), Tx("dlg.pickFolderBatch", UCase$(FmtExt(fmt))))
        If Len(folder) = 0 Then Exit Sub
    End If

    ResetRunState
    baseName = SafeName(BaseNameOf(pres.Name))
    If gBusy Then Exit Sub
    gBusy = True
    ProgressBegin Tx("prog.batch", slideList.Count, UCase$(FmtExt(fmt))), slideList.Count
    For i = 1 To slideList.Count
        Set sld = slideList(i)
        ProgressItem i, Tx("prog.page", i, slideList.Count, sld.SlideIndex)
        CheckCancel
        Set rng = ExportableRange(sld)
        If rng Is Nothing Then
            skipped = skipped + 1
        Else
            path = UniquePath(folder, baseName & "_" & Tx("file.page", sld.SlideIndex), FmtExt(fmt))
            p = ExportItem(pres, rng, path, fmt)
            If Len(p) > 0 Then files.Add p
        End If
    Next i
    ProgressEnd
    gBusy = False

    If files.Count = 0 Then
        MsgBox Tx("msg.noShapesSlides"), vbInformation, AppTitle()
        Exit Sub
    End If
    If skipped > 0 Then
        MsgBox Tx("msg.skipped", skipped), vbInformation, AppTitle()
    End If
    Report files, True
    Exit Sub
EH:
    ProgressEnd
    gBusy = False
    If Err.Number = ERR_CANCEL Then
        ReportCancelled files
    Else
        MsgBox Tx("msg.batchFailed", Err.Description), vbExclamation, AppTitle()
    End If
End Sub

' Pastes an EMF picture copy of the selection next to it (the original stays untouched).
Private Sub DoPictureCopy()
    Dim win As Object, sr As Object, sld As Object, pic As Object
    On Error GoTo EH
    If Not GetTarget(win, sr, False) Then Exit Sub
    Set sld = sr.Parent
    Set pic = PasteFromClipboard(sld, sr, True).Item(1)
    pic.Name = "VecStamp EMF " & Format$(Now, "hhnnss")
    pic.Select
    Exit Sub
EH:
    MsgBox Tx("msg.pictureFailed", Err.Description), vbExclamation, AppTitle()
End Sub

' The system colour dialog for the custom background (Windows: the common colour dialog,
' macOS: the macOS colour panel). Public macro VS_PickColor.
Private Sub DoPickColor()
    Dim c As Long
    On Error GoTo EH
    c = BgColorValue(BG_CUSTOM)
#If Mac Then
    If MacColorPicker(c) Then SetBgColor c
#Else
    If SysColorDialog(c, GetActiveWindow()) Then SetBgColor c
#End If
    Exit Sub
EH:
    MsgBox Tx("msg.colorFailed", Err.Description), vbExclamation, AppTitle()
End Sub

#If Mac Then
' macOS colour panel (the same panel PowerPoint's "More Colors" opens on the Mac). Without the
' helper script VecStamp's own colour list is used instead.
Private Function MacColorPicker(ByRef c As Long) As Boolean
    Dim res As String, r As Long, g As Long, b As Long
    On Error GoTo Fallback
    res = MacScript2("chooseColor", (c And &HFF&) & "," & ((c \ &H100&) And &HFF&) & "," & ((c \ &H10000) And &HFF&))
    If res = "error" Then GoTo Fallback
    If Len(res) = 0 Then Exit Function                  ' cancelled
    If ParseColor(res, r, g, b) Then
        c = RGB(r, g, b)
        MacColorPicker = True
    End If
    Exit Function
Fallback:
    MacColorPicker = ColorListPanel(c)
End Function
#Else
Private Function SysColorDialog(ByRef c As Long, ByVal owner As LongPtr) As Boolean
    Dim cc As CHOOSECOLORW
    cc.lStructSize = LenB(cc)
    cc.hwndOwner = owner
    cc.rgbResult = c
    cc.lpCustColors = VarPtr(gCustColors(0))
    cc.Flags = CC_RGBINIT Or CC_FULLOPEN Or CC_ANYCOLOR
    If ChooseColorW(cc) = 0 Then Exit Function
    c = cc.rgbResult And &HFFFFFF
    SysColorDialog = True
End Function

' VecStamp's own colour panel, used when PowerPoint's "More Colors" dialog cannot be opened:
' theme colours with tints and shades, standard and recent colours, a colour value field
' (R,G,B or #RRGGBB) with a preview, the system colour dialog, OK and Cancel. It is a native
' window driven by the message loop below - clicks are read from the messages, so no window
' procedure is needed. Returns 1 (c = chosen colour), 0 (cancelled) or -1 (no window).
Private Function ColorPanel(ByRef c As Long) As Long
    Dim owner As LongPtr, hwnd As LongPtr, hdc As LongPtr, font As LongPtr, fontB As LongPtr
    Dim hTheme As LongPtr, hStd As LongPtr, hRecent As LongPtr, hEdit As LongPtr, hPrev As LongPtr, hInfo As LongPtr
    Dim hOk As LongPtr, hCancel As LongPtr, hSys As LongPtr, clicked As LongPtr, f As LongPtr
    Dim bmp(0 To 2) As LongPtr, i As Long
    Dim m As MSG_T, r As RECT, k As Double, dpi As Long, cs As Long, gp As Long, rowGap As Long, lm As Long
    Dim x As Long, y As Long, w As Long, h As Long, cw As Long, ch As Long, px As Long, py As Long, pw As Long, ph As Long
    Dim idx As Long, act As Long, cur As Long, txt As String, lastTxt As String, handled As Boolean
    Dim rc As Variant, nRecent As Long, lastClick As Double, lastIdx As Long, lbl As String
    Dim rr As Long, gg As Long, bb As Long, cls As String, cap As String, face As String, cols() As Long

    On Error GoTo Failed
    ColorPanel = -1
    owner = GetActiveWindow()
    hdc = GetDC(0)
    dpi = GetDeviceCaps(hdc, LOGPIXELSY)
    ReleaseDC 0, hdc
    If dpi < 72 Then dpi = 96
    k = dpi / 96#
    cs = CLng(18 * k): gp = CLng(4 * k): rowGap = CLng(5 * k): lm = CLng(16 * k)
    rc = RecentColors()
    nRecent = UBound(rc) + 1
    cur = c
    lastIdx = -1

    ' layout (client coordinates)
    cw = 2 * lm + 10 * cs + 9 * gp
    If cw < CLng(320 * k) Then cw = CLng(320 * k)
    pw = CLng(64 * k): ph = CLng(24 * k)
    y = CLng(12 * k)
    ch = y + CLng(20 * k) + 6 * cs + rowGap + CLng(10 * k) + CLng(20 * k) + cs + CLng(10 * k)
    If nRecent > 0 Then ch = ch + CLng(20 * k) + cs + CLng(10 * k)
    ch = ch + CLng(32 * k) + CLng(42 * k) + CLng(42 * k)

    r.Left = 0: r.Top = 0: r.Right = cw: r.Bottom = ch
    AdjustWindowRectEx r, WS_POPUP Or WS_CAPTION, 0, WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT
    w = r.Right - r.Left
    h = r.Bottom - r.Top
    If owner <> 0 And GetWindowRect(owner, r) <> 0 Then
        x = r.Left + (r.Right - r.Left - w) \ 2
        y = r.Top + (r.Bottom - r.Top - h) \ 3
    Else
        x = (GetSystemMetrics(0) - w) \ 2
        y = (GetSystemMetrics(1) - h) \ 3
    End If
    cls = "#32770"
    cap = AppTitle() & " - " & Tx("col.title")
    hwnd = CreateWindowExW(WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT, StrPtr(cls), StrPtr(cap), _
                           WS_POPUP Or WS_CAPTION, x, y, w, h, owner, 0, 0, 0)
    If hwnd = 0 Then Exit Function
    face = UiFontName()
    font = CreateFontW(-CLng(9 * dpi / 72), 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 5, 0, StrPtr(face))
    fontB = CreateFontW(-CLng(9 * dpi / 72), 0, 0, 0, 700, 0, 0, 0, 1, 0, 0, 5, 0, StrPtr(face))

    ' theme colours (row 0) with their tints and shades (rows 1-5)
    y = CLng(12 * k)
    PanelChild hwnd, "Static", Tx("col.theme"), SS_NOPREFIX, lm, y, cw - 2 * lm, CLng(18 * k), fontB
    y = y + CLng(20 * k)
    ReDim cols(0 To 59)
    For i = 0 To 59: cols(i) = GalColor(i): Next i
    hTheme = PanelChild(hwnd, "Static", "", SS_BITMAP Or SS_NOTIFY, lm, y, 0, 0, 0)
    bmp(0) = PanelBitmap(hTheme, "theme", 10, 6, cols, cs, cs, gp, rowGap)
    y = y + 6 * cs + rowGap + CLng(10 * k)
    ' standard colours
    PanelChild hwnd, "Static", Tx("col.standard"), SS_NOPREFIX, lm, y, cw - 2 * lm, CLng(18 * k), fontB
    y = y + CLng(20 * k)
    ReDim cols(0 To 9)
    For i = 0 To 9: cols(i) = GalColor(60 + i): Next i
    hStd = PanelChild(hwnd, "Static", "", SS_BITMAP Or SS_NOTIFY, lm, y, 0, 0, 0)
    bmp(1) = PanelBitmap(hStd, "std", 10, 1, cols, cs, cs, gp, 0)
    y = y + cs + CLng(10 * k)
    ' recent colours
    If nRecent > 0 Then
        PanelChild hwnd, "Static", Tx("col.recent"), SS_NOPREFIX, lm, y, cw - 2 * lm, CLng(18 * k), fontB
        y = y + CLng(20 * k)
        ReDim cols(0 To nRecent - 1)
        For i = 0 To nRecent - 1: cols(i) = CLng(rc(i)): Next i
        hRecent = PanelChild(hwnd, "Static", "", SS_BITMAP Or SS_NOTIFY, lm, y, 0, 0, 0)
        bmp(2) = PanelBitmap(hRecent, "recent", nRecent, 1, cols, cs, cs, gp, 0)
        y = y + cs + CLng(10 * k)
    End If
    ' colour value, preview and the name of the colour
    PanelChild hwnd, "Static", Tx("col.value"), SS_NOPREFIX, lm, y + CLng(4 * k), CLng(80 * k), CLng(18 * k), font
    hEdit = PanelChild(hwnd, "Edit", HexRGB(cur), WS_TABSTOP Or ES_AUTOHSCROLL, lm + CLng(84 * k), y, _
                       CLng(110 * k), ph, font, WS_EX_CLIENTEDGE)
    SendMessageW hEdit, EM_LIMITTEXT, 24, 0
    hPrev = PanelChild(hwnd, "Static", "", SS_BITMAP, lm + CLng(202 * k), y, 0, 0, 0)
    PanelPreview hPrev, cur, pw, ph
    y = y + CLng(32 * k)
    hInfo = PanelChild(hwnd, "Static", Tx("col.hint"), SS_NOPREFIX, lm, y, cw - 2 * lm, CLng(36 * k), font)
    y = y + CLng(42 * k)
    hSys = PanelChild(hwnd, "Button", Tx("col.system"), WS_TABSTOP, lm, y, CLng(128 * k), CLng(28 * k), font)
    hOk = PanelChild(hwnd, "Button", Tx("btn.ok"), WS_TABSTOP Or BS_DEFPUSHBUTTON, cw - lm - CLng(176 * k), y, _
                     CLng(84 * k), CLng(28 * k), font)
    hCancel = PanelChild(hwnd, "Button", Tx("btn.cancel"), WS_TABSTOP, cw - lm - CLng(84 * k), y, _
                         CLng(84 * k), CLng(28 * k), font)

    ' modal loop
    If owner <> 0 Then EnableWindow owner, 0
    ShowWindow hwnd, SW_SHOW
    UpdateWindow hwnd
    SetFocusW hEdit
    SendMessageW hEdit, EM_SETSEL, 0, -1
    lastTxt = HexRGB(cur)
    ColorPanel = 0
    Do
        If IsWindow(hwnd) = 0 Then Exit Do
        If PeekMessageW(m, 0, 0, 0, PM_REMOVE) <> 0 Then
            handled = False
            Select Case m.message
                Case WM_QUIT
                    PostQuitMessage CLng(m.wParam)
                    act = 2
                    handled = True
                Case WM_LBUTTONDOWN, WM_LBUTTONDBLCLK
                    px = CLng(m.lParam And &HFFFF&)
                    py = CLng((m.lParam \ &H10000) And &HFFFF&)
                    idx = -1
                    If m.hwnd = hTheme And hTheme <> 0 Then idx = PanelHit(px, py, 10, 6, cs, gp, rowGap, 0)
                    If m.hwnd = hStd And hStd <> 0 Then idx = PanelHit(px, py, 10, 1, cs, gp, 0, 60)
                    If m.hwnd = hRecent And hRecent <> 0 Then idx = PanelHit(px, py, nRecent, 1, cs, gp, 0, 70)
                    If idx >= 0 Then
                        cur = GalColor(idx, lbl)
                        lastTxt = HexRGB(cur)
                        SetWindowTextW hEdit, StrPtr(lastTxt)
                        PanelPreview hPrev, cur, pw, ph
                        SetWindowTextW hInfo, StrPtr(lbl)
                        If m.message = WM_LBUTTONDBLCLK Or (idx = lastIdx And Abs(Timer - lastClick) < GetDoubleClickTime() / 1000#) Then act = 1
                        lastIdx = idx
                        lastClick = Timer
                        handled = True
                    End If
                Case WM_LBUTTONUP
                    If m.hwnd = hOk Or m.hwnd = hCancel Or m.hwnd = hSys Then
                        GetClientRect m.hwnd, r
                        px = CLng(m.lParam And &HFFFF&)
                        py = CLng((m.lParam \ &H10000) And &HFFFF&)
                        If px < r.Right And py < r.Bottom Then clicked = m.hwnd
                    End If
                Case WM_KEYDOWN
                    Select Case CLng(m.wParam)
                        Case VK_ESCAPE
                            act = 2
                            handled = True
                        Case VK_RETURN
                            f = GetFocus()
                            If f = hCancel Then
                                act = 2
                            ElseIf f = hSys Then
                                clicked = hSys
                            Else
                                act = 1
                            End If
                            handled = True
                    End Select
                Case WM_KEYUP
                    If CLng(m.wParam) = VK_SPACE Then
                        f = GetFocus()
                        If f = hOk Or f = hCancel Or f = hSys Then clicked = f
                    End If
            End Select
            If Not handled Then
                If IsDialogMessageW(hwnd, m) = 0 Then
                    TranslateMessage m
                    DispatchMessageW m
                End If
            End If
            If clicked <> 0 Then
                If clicked = hOk Then act = 1
                If clicked = hCancel Then act = 2
                If clicked = hSys Then
                    If SysColorDialog(cur, hwnd) Then
                        lastTxt = HexRGB(cur)
                        SetWindowTextW hEdit, StrPtr(lastTxt)
                        PanelPreview hPrev, cur, pw, ph
                        lbl = Tx("col.withValue", Tx("col.systemPicked"), ColorText(cur))
                        SetWindowTextW hInfo, StrPtr(lbl)
                    End If
                End If
                clicked = 0
            End If
        Else
            SleepX 10
        End If
        txt = WindowText(hEdit)
        If txt <> lastTxt Then                      ' typed colour value
            lastTxt = txt
            If ParseColor(txt, rr, gg, bb) Then
                cur = RGB(rr, gg, bb)
                PanelPreview hPrev, cur, pw, ph
                SetWindowTextW hInfo, StrPtr(ColorText(cur))
            End If
        End If
        If act = 1 Then
            If ParseColor(WindowText(hEdit), rr, gg, bb) Then
                c = RGB(rr, gg, bb)
                ColorPanel = 1
                Exit Do
            End If
            lbl = Tx("col.bad")
            SetWindowTextW hInfo, StrPtr(lbl)
            act = 0
        ElseIf act = 2 Then
            Exit Do
        End If
    Loop

Cleanup:
    On Error Resume Next
    If owner <> 0 Then EnableWindow owner, 1
    If hPrev <> 0 Then DeleteObject SendMessageW(hPrev, STM_SETIMAGE, IMAGE_BITMAP, 0)
    If hwnd <> 0 Then DestroyWindow hwnd
    For i = 0 To 2
        If bmp(i) <> 0 Then DeleteObject bmp(i)
    Next i
    If font <> 0 Then DeleteObject font
    If fontB <> 0 Then DeleteObject fontB
    Exit Function
Failed:
    If hwnd = 0 Then ColorPanel = -1 Else ColorPanel = 0
    Resume Cleanup
End Function

Private Function PanelChild(ByVal parent As LongPtr, ByVal cls As String, ByVal text As String, ByVal style As Long, _
                            ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal font As LongPtr, _
                            Optional ByVal exStyle As Long = 0) As LongPtr
    Dim hw As LongPtr
    hw = CreateWindowExW(exStyle, StrPtr(cls), StrPtr(text), WS_CHILD Or WS_VISIBLE Or style, x, y, w, h, parent, 0, 0, 0)
    If hw <> 0 And font <> 0 Then SendMessageW hw, WM_SETFONT, font, 1
    PanelChild = hw
End Function

Private Function WindowText(ByVal hw As LongPtr) As String
    Dim buf As String, n As Long
    buf = String$(128, vbNullChar)
    n = GetWindowTextW(hw, StrPtr(buf), 128)
    If n > 0 Then WindowText = Left$(buf, n)
End Function

' Index of the colour cell at (x, y) of a palette bitmap (first = index of its first cell), or -1.
Private Function PanelHit(ByVal x As Long, ByVal y As Long, ByVal cols As Long, ByVal rows As Long, ByVal cs As Long, _
                          ByVal gp As Long, ByVal rowGap As Long, ByVal first As Long) As Long
    Dim col As Long, row As Long
    PanelHit = -1
    If x Mod (cs + gp) >= cs Then Exit Function
    col = x \ (cs + gp)
    If col >= cols Then Exit Function
    If y < cs Then
        row = 0
    ElseIf rows > 1 And y >= cs + rowGap Then
        row = 1 + (y - cs - rowGap) \ cs
    Else
        Exit Function
    End If
    If row >= rows Then Exit Function
    PanelHit = first + row * cols + col
End Function

' Shows a palette (cols x rows colour cells) in a bitmap static control; returns the bitmap.
Private Function PanelBitmap(ByVal hStatic As LongPtr, ByVal name As String, ByVal cols As Long, ByVal rows As Long, _
                             ByRef colors() As Long, ByVal cellW As Long, ByVal cellH As Long, ByVal gp As Long, _
                             ByVal rowGap As Long) As LongPtr
    Dim p As String, hb As LongPtr
    p = JoinPath(TempDirX(), "panel_" & name & "_" & cellW & ".bmp")
    CellsBmp p, cols, rows, colors, cellW, cellH, gp, rowGap
    hb = LoadImageW(0, StrPtr(p), IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
    If hb <> 0 Then SendMessageW hStatic, STM_SETIMAGE, IMAGE_BITMAP, hb
    PanelBitmap = hb
End Function

Private Sub PanelPreview(ByVal hStatic As LongPtr, ByVal c As Long, ByVal w As Long, ByVal h As Long)
    Dim p As String, hb As LongPtr, old As LongPtr, one(0 To 0) As Long
    one(0) = c
    p = JoinPath(TempDirX(), "panel_preview_" & w & ".bmp")
    CellsBmp p, 1, 1, one, w, h, 0, 0
    hb = LoadImageW(0, StrPtr(p), IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
    If hb = 0 Then Exit Sub
    old = SendMessageW(hStatic, STM_SETIMAGE, IMAGE_BITMAP, hb)
    If old <> 0 Then DeleteObject old
End Sub

' Writes a 24-bit BMP of colour cells (row 0 is separated from the other rows by rowGap) on the
' dialog background colour; every cell has a grey frame.
Private Sub CellsBmp(ByVal p As String, ByVal cols As Long, ByVal rows As Long, ByRef colors() As Long, _
                     ByVal cellW As Long, ByVal cellH As Long, ByVal gp As Long, ByVal rowGap As Long)
    Dim b() As Byte, w As Long, h As Long, rowBytes As Long, n As Long, x As Long, y As Long, i As Long
    Dim col As Long, row As Long, x0 As Long, y0 As Long, c As Long, bg As Long, raw As String
    w = cols * cellW + (cols - 1) * gp
    h = rows * cellH
    If rows > 1 Then h = h + rowGap
    rowBytes = ((w * 3 + 3) \ 4) * 4
    n = 54 + rowBytes * h
    ReDim b(0 To n - 1)
    BmpHeader b, w, h, rowBytes
    bg = GetSysColor(COLOR_BTNFACE)
    For y = 0 To h - 1
        For x = 0 To w - 1
            i = 54 + y * rowBytes + x * 3
            b(i) = (bg \ &H10000) And &HFF&: b(i + 1) = (bg \ &H100&) And &HFF&: b(i + 2) = bg And &HFF&
        Next x
    Next y
    For row = 0 To rows - 1
        For col = 0 To cols - 1
            c = colors(row * cols + col)
            x0 = col * (cellW + gp)
            y0 = row * cellH
            If row > 0 Then y0 = y0 + rowGap
            For y = y0 To y0 + cellH - 1
                For x = x0 To x0 + cellW - 1
                    i = 54 + (h - 1 - y) * rowBytes + x * 3
                    If x = x0 Or y = y0 Or x = x0 + cellW - 1 Or y = y0 + cellH - 1 Then
                        b(i) = 170: b(i + 1) = 170: b(i + 2) = 170
                    Else
                        b(i) = (c \ &H10000) And &HFF&: b(i + 1) = (c \ &H100&) And &HFF&: b(i + 2) = c And &HFF&
                    End If
                Next x
            Next y
        Next col
    Next row
    raw = b
    WriteBytesX p, raw
End Sub
#End If

Private Sub BmpHeader(ByRef b() As Byte, ByVal w As Long, ByVal h As Long, ByVal rowBytes As Long)
    b(0) = 66: b(1) = 77                                         ' "BM"
    PutLong b, 2, 54 + rowBytes * h
    PutLong b, 10, 54
    PutLong b, 14, 40
    PutLong b, 18, w
    PutLong b, 22, h
    b(26) = 1: b(28) = 24
    PutLong b, 34, rowBytes * h
End Sub

' Last-resort colour panel when no native window can be shown: a numbered list of theme,
' standard and recent colours in an input box; a number, R,G,B or #RRGGBB can be entered.
Private Function ColorListPanel(ByRef c As Long) As Boolean
    Dim msg As String, ans As String, i As Long, n As Long, lbl As String, idx(0 To 29) As Long
    Dim rc As Variant, r As Long, g As Long, b As Long
    For i = 0 To 9
        idx(n) = i: n = n + 1                                    ' theme colours
    Next i
    For i = 0 To 9
        idx(n) = 60 + i: n = n + 1                               ' standard colours
    Next i
    rc = RecentColors()
    For i = 0 To UBound(rc)
        If i < 5 Then idx(n) = 70 + i: n = n + 1
    Next i
    msg = Tx("col.listPrompt") & vbCrLf
    For i = 0 To n - 1
        GalColor idx(i), lbl
        msg = msg & (i + 1) & ". " & lbl & vbCrLf
    Next i
    Do
        ans = Trim$(InputBox(msg, AppTitle() & " - " & Tx("col.title"), HexRGB(c)))
        If Len(ans) = 0 Then Exit Function
        If IsNumeric(ans) And InStr(ans, ",") = 0 Then
            If CLng(Val(ans)) >= 1 And CLng(Val(ans)) <= n Then
                c = GalColor(idx(CLng(Val(ans)) - 1))
                ColorListPanel = True
                Exit Function
            End If
        ElseIf ParseColor(ans, r, g, b) Then
            c = RGB(r, g, b)
            ColorListPanel = True
            Exit Function
        End If
        MsgBox Tx("col.bad"), vbExclamation, AppTitle()
    Loop
End Function

Private Sub ResetRunState()
    gSvgFallback = 0
    gSvgNative = 0
    gSvgXpsOff = False
    gBigPngOk = False
End Sub

' What to export: the selection, or (after asking) every shape on the current slide.
Private Function GetTarget(ByRef win As Object, ByRef sr As Object, ByVal allowWholeSlide As Boolean) As Boolean
    Dim sld As Object, selType As Long
    On Error Resume Next
    Set win = Application.ActiveWindow
    selType = -1
    selType = win.Selection.Type
    On Error GoTo 0
    If win Is Nothing Or selType < 0 Then
        MsgBox Tx("msg.openAndSelect"), vbInformation, AppTitle()
        Exit Function
    End If

    If selType = PP_SEL_SHAPES Or selType = PP_SEL_TEXT Then
        If win.Selection.HasChildShapeRange Then
            Set sr = win.Selection.ChildShapeRange
        Else
            Set sr = win.Selection.ShapeRange
        End If
    Else
        If Not allowWholeSlide Then
            MsgBox Tx("msg.selectFirst"), vbInformation, AppTitle()
            Exit Function
        End If
        On Error Resume Next
        Set sld = win.View.Slide
        On Error GoTo 0
        If sld Is Nothing Then
            MsgBox Tx("msg.normalView"), vbInformation, AppTitle()
            Exit Function
        End If
        Set sr = ExportableRange(sld)
        If sr Is Nothing Then
            MsgBox Tx("msg.noShapesSlide"), vbInformation, AppTitle()
            Exit Function
        End If
        If MsgBox(Tx("msg.askWholeSlide"), vbQuestion + vbYesNo, AppTitle()) <> vbYes Then Exit Function
    End If
    GetTarget = Not (sr Is Nothing)
End Function

' All visible shapes on a slide except empty placeholders ("click to add title" boxes).
Private Function ExportableRange(ByVal sld As Object) As Object
    Dim i As Long, n As Long, shp As Object, keep As Boolean
    Dim idx() As Variant
    If sld.Shapes.Count = 0 Then Exit Function
    ReDim idx(1 To sld.Shapes.Count)
    For i = 1 To sld.Shapes.Count
        Set shp = sld.Shapes.Item(i)
        keep = True
        On Error Resume Next
        If shp.Visible = MSO_FALSE Then keep = False
        If shp.Type = MSO_PLACEHOLDER Then
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText = MSO_FALSE Then keep = False
            End If
        End If
        On Error GoTo 0
        If keep Then
            n = n + 1
            idx(n) = i
        End If
    Next i
    If n = 0 Then Exit Function
    ReDim Preserve idx(1 To n)
    Set ExportableRange = sld.Shapes.Range(idx)
End Function

' Exports one shape or a range of shapes. Returns the file written ("" if skipped).
' Several shapes, a background or a margin are composed on a temporary slide.
Private Function ExportItem(ByVal pres As Object, ByVal src As Object, _
                            ByVal path As String, ByVal fmt As Long) As String
    Dim n As Long, margin As Single
    Dim tmp As Object, target As Object, wasSaved As Long
    Dim errNum As Long, errDesc As String, result As String

    n = 1
    If TypeName(src) = "ShapeRange" Then n = src.Count
    margin = GetMargin()
    ProgressAt 0.03, Tx("step.prepare", n)
    CheckCancel

    If n = 1 And BgTypeIdx() = BG_NONE And margin <= 0 Then
        If TypeName(src) = "ShapeRange" Then
            ExportItem = ExportShape(pres, src.Item(1), path, fmt)
        Else
            ExportItem = ExportShape(pres, src, path, fmt)
        End If
        Exit Function
    End If

    wasSaved = pres.Saved
    On Error GoTo Failed
    ProgressAt 0.08, Tx("step.compose")
    Set tmp = AddTempSlide(pres, src)
    Set target = BuildComposite(tmp, src, margin)
    ProgressAt 0.18, Tx("step.composed")
    CheckCancel
    result = ExportShape(pres, target, path, fmt)

Cleanup:
    On Error Resume Next
    If Not tmp Is Nothing Then tmp.Delete
    pres.Saved = wasSaved
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, , errDesc
    ExportItem = result
    Exit Function
Failed:
    errNum = Err.Number
    errDesc = Err.Description
    Resume Cleanup
End Function

' Writes one shape in the requested format. Returns the path written ("" if skipped).
Private Function ExportShape(ByVal pres As Object, ByVal shp As Object, _
                             ByVal path As String, ByVal fmt As Long) As String
    Dim t As String, dpi As Long, px As Double
    Select Case fmt
        Case FMT_EMF
            ProgressAt 0.25, Tx("step.ppEmf")
            ExportRaw shp, path, PP_FMT_EMF
        Case FMT_WMF
            ProgressAt 0.25, Tx("step.ppWmf")
            ExportRaw shp, path, PP_FMT_WMF
        Case FMT_EMZ
            ProgressAt 0.25, Tx("step.ppEmf")
            t = TempFileX("emf")
            ExportToTemp shp, t, PP_FMT_EMF, 0, 0
            ProgressAt 0.65, Tx("step.emz")
            If Not GzipX(t, path) Then Err.Raise vbObjectError + 514, , Tx("err.emz")
        Case FMT_PNG
            dpi = DpiValue()
            px = (shp.Width / 72# * dpi) * (shp.Height / 72# * dpi)
            If px > PNG_WARN_PIXELS And Not gBigPngOk Then
                If AskUser(Tx("msg.bigPng", dpi, CLng(shp.Width / 72# * dpi), CLng(shp.Height / 72# * dpi), _
                              Format$(px / 1000000#, "0")), vbQuestion + vbYesNo) <> vbYes Then
                    Err.Raise vbObjectError + 599, , "cancelled"
                End If
                gBigPngOk = True
            End If
            ProgressAt 0.25, Tx("step.png", CLng(shp.Width / 72# * dpi), CLng(shp.Height / 72# * dpi), dpi)
            ExportRaw shp, path, PP_FMT_PNG, _
                      CLng(pres.PageSetup.SlideWidth * dpi / 96), CLng(pres.PageSetup.SlideHeight * dpi / 96)
        Case FMT_SVG
            path = ExportSvg(pres, shp, path)
        Case FMT_PDF
            ExportPdf shp, path
        Case FMT_VSDX
            ProgressAt 0.2, Tx("step.analyze")
            ExportVsdx pres, shp, path
    End Select
    If Len(path) > 0 Then
        If Not FileExistsX(path) Then Err.Raise vbObjectError + 515, , Tx("err.noFile", path)
    End If
    ProgressAt 1, Tx("step.doneFile", FileNameOf(path))
    ExportShape = path
End Function

' SVG, best method first:
'  1. PowerPoint's own SVG export (2302+ / Mac 16.82+), checked and repaired;
'  2. Windows, any PowerPoint 2013+: PowerPoint prints the drawing to XPS and the built-in
'     converter turns it into SVG (paths, gradients, images, real text, embedded fonts);
'  3. EMF converted with Inkscape or LibreOffice, when installed;
'  4. otherwise offer EMF instead.
Private Function ExportSvg(ByVal pres As Object, ByVal shp As Object, ByVal path As String) As String
    Dim ok As Boolean, t As String, e As String, p2 As String, st As Long

    If gSvgNative <> 2 Then
        ProgressAt 0.22, Tx("step.ppSvg")
        t = TempFileX("svg")
        On Error Resume Next
        ExportToTemp shp, t, PP_FMT_SVG, 0, 0
        ok = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
        ProgressAt 0.8, Tx("step.fixSvg")
        If ok Then ok = FixSvgFile(t)
        If ok Then
            gSvgNative = 1
            MoveTempTo t, path
            ExportSvg = path
            Exit Function
        End If
        gSvgNative = 2                              ' this PowerPoint cannot write SVG
        DeleteTempX t
        gPgFrac = 0.2                               ' the next method starts over
    End If

    st = RUN_FAILED
    t = TempFileX("svg")
    If Not gSvgXpsOff Then st = XpsToSvg(pres, shp, t)
    If st = RUN_CANCELLED Then
        DeleteTempX t
        Err.Raise ERR_CANCEL, , "cancelled"
    End If
    ok = (st = 0)
    If ok Then
        ProgressAt 0.96, Tx("step.fixSvg")
        ok = FixSvgFile(t)
    End If
    If ok Then
        MoveTempTo t, path
        ExportSvg = path
        Exit Function
    End If
    DeleteTempX t
    If st = RUN_TIMEOUT Then
        If AskUser(Tx("msg.svgTimeout", SVG_TIMEOUT), vbQuestion + vbYesNo) <> vbYes Then
            Err.Raise ERR_CANCEL, , "cancelled"
        End If
        gSvgXpsOff = True                           ' do not wait again for the next shapes of this run
    End If

    ProgressAt 0.5, Tx("step.emf2svg")
    e = TempFileX("emf")
    t = TempFileX("svg")
    ExportToTemp shp, e, PP_FMT_EMF, 0, 0
    st = ConvertEmfToSvg(e, t)
    If st = RUN_CANCELLED Then
        DeleteTempX e
        DeleteTempX t
        Err.Raise ERR_CANCEL, , "cancelled"
    End If
    ok = (st = 0)
    If ok Then ok = FixSvgFile(t)
    DeleteTempX e
    If ok Then
        MoveTempTo t, path
        ExportSvg = path
        Exit Function
    End If
    DeleteTempX t

    If gSvgFallback = 0 Then
        If AskUser(Tx("msg.noSvg"), vbQuestion + vbYesNo) = vbYes Then
            gSvgFallback = 1
        Else
            gSvgFallback = 2
        End If
    End If
    If gSvgFallback = 1 Then
        p2 = UniquePath(ParentFolder(path), BaseNameOf(path), "emf")
        ExportRaw shp, p2, PP_FMT_EMF
        ExportSvg = p2
    End If
End Function

' Makes sure an exported file really is SVG that browsers display as a picture:
' the root element must be <svg> in the SVG namespace (missing xmlns is added).
' Returns False when the file is not SVG at all.
Private Function FixSvgFile(ByVal p As String) As Boolean
    Dim data As String, k As Long, pos As Long, tagEnd As Long, nm As String
    Dim tag As String, ins As String, nextB As String

    data = ReadBytesX(p)
    If LenB(data) < 5 Then Exit Function
    If AscB(MidB(data, 1, 1)) = &HFF And AscB(MidB(data, 2, 1)) = &HFE Then
        FixSvgFile = (InStr(1, data, "<svg", vbTextCompare) > 0)      ' UTF-16 text: leave untouched
        Exit Function
    End If

    pos = 1
    Do
        k = InStrB(pos, data, AsciiB("<"))
        If k = 0 Or k > 65536 Then Exit Function
        nextB = MidB(data, k + 1, 1)
        If nextB = AsciiB("?") Or nextB = AsciiB("!") Then
            If MidB(data, k, 4) = AsciiB("<!--") Then
                pos = InStrB(k, data, AsciiB("-->"))
            Else
                pos = InStrB(k, data, AsciiB(">"))
            End If
            If pos = 0 Then Exit Function
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop

    tagEnd = InStrB(k, data, AsciiB(">"))
    If tagEnd = 0 Then Exit Function
    tag = MidB(data, k, tagEnd - k + 1)
    nm = LCase$(TagNameOf(tag))
    If nm = "svg" Then
        If InStrB(1, tag, AsciiB("http://www.w3.org/2000/svg")) = 0 Then
            ins = ins & AsciiB(" xmlns=""http://www.w3.org/2000/svg""")
        End If
        If InStrB(1, data, AsciiB("xlink:")) > 0 And InStrB(1, tag, AsciiB("xmlns:xlink")) = 0 Then
            ins = ins & AsciiB(" xmlns:xlink=""http://www.w3.org/1999/xlink""")
        End If
        If LenB(ins) > 0 Then
            data = LeftB(data, k + 3) & ins & MidB(data, k + 4)
            WriteBytesX p, data
        End If
        FixSvgFile = True
    ElseIf Right$(nm, 4) = ":svg" Then
        FixSvgFile = True
    End If
End Function

' Element name of a start tag held as bytes, e.g. "<svg width=..." -> "svg".
Private Function TagNameOf(ByVal tagBytes As String) As String
    Dim i As Long, c As Long, s As String
    For i = 2 To LenB(tagBytes)
        c = AscB(MidB(tagBytes, i, 1))
        If c = 32 Or c = 9 Or c = 10 Or c = 13 Or c = 62 Or c = 47 Then Exit For
        s = s & Chr$(c)
    Next i
    TagNameOf = s
End Function

' Vector PDF whose page is the size of the drawing (PowerPoint pages are 1 to 56 inches,
' so very small or very large drawings are scaled to fit those limits).
Private Sub ExportPdf(ByVal shp As Object, ByVal path As String)
    Dim tp As Object, sld As Object, pic As Object, t As String
    Dim w As Single, h As Single, k As Single, pw As Single, ph As Single
    Dim errNum As Long, errDesc As String

    w = shp.Width
    h = shp.Height
    If w < 1 Then w = 1
    If h < 1 Then h = 1
    k = 1
    If w < 72 Or h < 72 Then k = 72 / MinS(w, h)
    If MaxS(w, h) * k > 4032 Then k = 4032 / MaxS(w, h)
    pw = MinS(MaxS(w * k, 72), 4032)
    ph = MinS(MaxS(h * k, 72), 4032)
    t = TempFileX("pdf")
    ProgressAt 0.25, Tx("step.pdfDoc")

    On Error GoTo Failed
    Set tp = Application.Presentations.Add(MSO_FALSE)
    tp.PageSetup.SlideWidth = pw
    tp.PageSetup.SlideHeight = ph
    Set sld = tp.Slides.Add(1, PP_LAYOUT_BLANK)
    ProgressAt 0.45, Tx("step.pdfPaste")
    Set pic = PasteFromClipboard(sld, shp, True).Item(1)
    pic.LockAspectRatio = MSO_FALSE
    pic.Width = w * k
    pic.Height = h * k
    pic.Left = (pw - w * k) / 2
    pic.Top = (ph - h * k) / 2
    ProgressAt 0.7, Tx("step.pdfWrite")
    tp.SaveAs t, PP_SAVE_AS_PDF

Cleanup:
    On Error Resume Next
    If Not tp Is Nothing Then
        tp.Saved = MSO_TRUE
        tp.Close
    End If
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, , Tx("err.pdf", errDesc)
    MoveTempTo t, path
    Exit Sub
Failed:
    errNum = Err.Number
    errDesc = Err.Description
    Resume Cleanup
End Sub

' Visio drawing. Default: native Visio shapes built from the PowerPoint shapes (geometry,
' fill, line, arrowheads, rotation, editable rich text, groups) - no Visio needed.
' Pictures, charts, tables and other objects without a Visio equivalent are embedded as
' EMF pictures. "Picture mode" (setting VsdxNative = 0) embeds the whole drawing as one EMF.
Private Sub ExportVsdx(ByVal pres As Object, ByVal shp As Object, ByVal path As String)
    Dim e As String
    If GetFlag("VsdxNative", "1") Then
        If VsdxNative(pres, shp, path) Then Exit Sub
    End If
    ProgressAt 0.3, Tx("step.vsdxPicture")
    e = TempFileX("emf")
    ExportToTemp shp, e, PP_FMT_EMF, 0, 0
    gVxId = 0
    Set gVxMedia = New Collection
    SbReset
    SbAdd VxForeignXml(e, 0, 0, shp.Width, shp.Height, "VecStamp")
    ProgressAt 0.7, Tx("step.vsdxWrite0")
    WriteVsdxPackage shp.Width / 72#, shp.Height / 72#, SbText(), path
    DeleteTempX e
End Sub

Private Function VsdxNative(ByVal pres As Object, ByVal shp As Object, ByVal path As String) As Boolean
    Dim wasSaved As Long, ok As Boolean, i As Long, errNum As Long
    wasSaved = pres.Saved
    Set gVxPres = pres
    gVxId = 0
    gVxNative = 0
    gVxPictures = 0
    gVxCopy = 0
    gVxDone = 0
    gVxTotal = VxCount(shp)
    If gVxTotal < 1 Then gVxTotal = 1
    Set gVxMedia = New Collection
    SbReset
    On Error GoTo Failed
    Set gVxScratch = AddTempSlide(pres, shp)
    VxShape shp, shp.Left, shp.Top + shp.Height
    If gVxNative = 0 And gVxPictures = 0 Then GoTo Cleanup
    ProgressAt 0.88, Tx("step.vsdxWrite", gVxNative, gVxPictures)
    WriteVsdxPackage shp.Width / 72#, shp.Height / 72#, SbText(), path
    ok = True
Cleanup:
    On Error Resume Next
    If Not gVxScratch Is Nothing Then gVxScratch.Delete
    Set gVxScratch = Nothing
    For i = 1 To gVxMedia.Count
        DeleteTempX CStr(gVxMedia(i))
    Next i
    pres.Saved = wasSaved
    On Error GoTo 0
    If errNum = ERR_CANCEL Then Err.Raise ERR_CANCEL, , "cancelled"
    VsdxNative = ok
    Exit Function
Failed:
    ok = False
    errNum = Err.Number
    Resume Cleanup
End Function

' Number of shapes the walker will convert (group members counted one by one).
Private Function VxCount(ByVal s As Object) As Long
    Dim i As Long, t As Long
    On Error Resume Next
    t = s.Type
    If t = MSO_GROUP Then
        For i = 1 To s.GroupItems.Count
            VxCount = VxCount + 1
        Next i
    Else
        VxCount = 1
    End If
End Function

' Writes the package: template parts + shapes XML + embedded EMF pictures, zipped.
Private Sub WriteVsdxPackage(ByVal wIn As Double, ByVal hIn As Double, ByVal shapesXml As String, ByVal path As String)
    Dim d As String, z As String, m As Variant, i As Long, parts() As String, txt As String
    Dim dst As String, sub1 As Variant, rels As String

    d = JoinPath(TempDirX(), "vsdx_" & Format$(Now, "hhnnss") & "_" & CStr(Int(Rnd * 1000000)))
    MakeDirX d
    For Each sub1 In Array("_rels", "docProps", "visio", "visio/_rels", "visio/pages", "visio/pages/_rels", "visio/media")
        MakeDirX JoinPath(d, Replace(CStr(sub1), "/", PathSep()))
    Next
    For i = 1 To gVxMedia.Count
        rels = rels & "<Relationship Id=""rId" & i & """ Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"" Target=""../media/image" & i & ".emf""/>"
        CopyFileX CStr(gVxMedia(i)), JoinPath(d, Replace("visio/media/image" & i & ".emf", "/", PathSep()))
    Next i
    m = VsdxMap()
    For i = LBound(m) To UBound(m)
        parts = Split(CStr(m(i)), "=")
        txt = VsdxPart(parts(0))
        txt = Replace(txt, "{PW}", NumXml(wIn))
        txt = Replace(txt, "{PH}", NumXml(hIn))
        txt = Replace(txt, "{W}", NumXml(wIn))
        txt = Replace(txt, "{H}", NumXml(hIn))
        txt = Replace(txt, "{CX}", NumXml(wIn / 2))
        txt = Replace(txt, "{CY}", NumXml(hIn / 2))
        txt = Replace(txt, "{LX}", NumXml(wIn / 2))
        txt = Replace(txt, "{LY}", NumXml(hIn / 2))
        txt = Replace(txt, "{VERSION}", VS_VERSION)
        txt = Replace(txt, "{TITLE}", "VecStamp drawing")
        txt = Replace(txt, "{CREATED}", Format$(Now, "yyyy-mm-dd") & "T" & Format$(Now, "hh:nn:ss") & "Z")
        txt = Replace(txt, "{RELS}", rels)
        txt = Replace(txt, "{SHAPES}", shapesXml)     ' last: shape text may contain braces
        dst = JoinPath(d, Replace(parts(1), "/", PathSep()))
        WriteBytesX dst, Utf8Bytes(txt)
    Next i
    z = TempFileX("vsdx")
    ZipFolderX d, z
    RemoveTempDirX d
    MoveTempTo z, path
End Sub

' ---- shape walker -------------------------------------------------------------
' ox, oy: origin of the parent (page or group) in slide points: left and bottom edge.
Private Sub VxShape(ByVal s As Object, ByVal ox As Double, ByVal oy As Double)
    Dim t As Long, i As Long, x As String, nm As String
    t = -1
    On Error Resume Next
    t = s.Type
    If s.Visible = MSO_FALSE Then Exit Sub
    On Error GoTo 0
    If t = MSO_GROUP Then
        gVxId = gVxId + 1
        nm = VxName(s)
        x = "<Shape ID='" & gVxId & "' NameU='" & nm & "' Name='" & nm & "' Type='Group' LineStyle='0' FillStyle='0' TextStyle='0'>" & _
            VxXform(s.Left - ox, oy - (s.Top + s.Height), s.Width, s.Height, s.Rotation) & "<Shapes>"
        SbAdd x
        For i = 1 To s.GroupItems.Count
            VxShape s.GroupItems.Item(i), s.Left, s.Top + s.Height
        Next i
        SbAdd "</Shapes></Shape>"
        Exit Sub
    End If
    gVxDone = gVxDone + 1
    ProgressAt 0.22 + 0.64 * (gVxDone - 1) / gVxTotal, Tx("step.vsdxShape", gVxDone, gVxTotal, ShortName(VxName(s)))
    CheckCancel
    If Not VxNative(s, t, ox, oy) Then VxPicture s, ox, oy
End Sub

Private Function VxNative(ByVal s As Object, ByVal t As Long, ByVal ox As Double, ByVal oy As Double) As Boolean
    Dim w As Double, h As Double, x As String, cells As String, sects As String, txt As String, nm As String
    Select Case t
        Case MSO_AUTOSHAPE, MSO_CALLOUT, MSO_FREEFORM, MSO_LINE, MSO_TEXTBOX, MSO_TEXTEFFECT
        Case MSO_PLACEHOLDER
            If Not PlaceholderIsText(s) Then Exit Function
        Case Else
            Exit Function
    End Select
    On Error GoTo Nope
    w = s.Width
    h = s.Height
    gGeoTextBox = False
    If Not VxGeometry(s, t, w, h) Then Exit Function
    txt = VxText(s, w, h, cells, sects)
    nm = VxName(s)
    x = "<Shape ID='" & (gVxId + 1) & "' NameU='" & nm & "' Name='" & nm & "' Type='Shape' LineStyle='0' FillStyle='0' TextStyle='0'>" & _
        VxXform(s.Left - ox, oy - (s.Top + h), w, h, s.Rotation) & VxStyle(s, t) & cells & VxGradient(s, t) & _
        sects & gGeo & txt & "</Shape>"
    gVxId = gVxId + 1
    SbAdd x
    gVxNative = gVxNative + 1
    VxNative = True
    Exit Function
Nope:
    VxNative = False
End Function

' Objects without a Visio equivalent: an EMF picture of the object, as a Visio foreign shape.
Private Sub VxPicture(ByVal s As Object, ByVal ox As Double, ByVal oy As Double)
    Dim e As String, c As Object, w As Double, h As Double, bw As Double, bh As Double, a As Double
    e = TempFileX("emf")
    On Error Resume Next
    s.Export e, PP_FMT_EMF
    If Err.Number <> 0 Then
        Err.Clear
        Set c = VxTopCopy(s)
        If Not c Is Nothing Then
            c.Export e, PP_FMT_EMF
            c.Delete
        End If
    End If
    On Error GoTo 0
    If Not FileExistsX(e) Then Exit Sub
    w = s.Width
    h = s.Height
    a = s.Rotation * 3.14159265358979 / 180#
    bw = Abs(w * Cos(a)) + Abs(h * Sin(a))          ' the picture shows the rotated object
    bh = Abs(w * Sin(a)) + Abs(h * Cos(a))
    SbAdd VxForeignXml(e, s.Left + w / 2 - bw / 2 - ox, oy - (s.Top + h / 2 + bh / 2), bw, bh, VxName(s))
    gVxPictures = gVxPictures + 1
End Sub

Private Function VxForeignXml(ByVal emfPath As String, ByVal xPt As Double, ByVal yPt As Double, _
                              ByVal wPt As Double, ByVal hPt As Double, ByVal nm As String) As String
    Dim w As String, h As String
    gVxMedia.Add emfPath
    gVxId = gVxId + 1
    w = NumXml(wPt / 72#)
    h = NumXml(hPt / 72#)
    VxForeignXml = "<Shape ID='" & gVxId & "' NameU='" & XmlAttr(nm) & "' Name='" & XmlAttr(nm) & "' Type='Foreign' LineStyle='0' FillStyle='0' TextStyle='0'>" & _
        VxXform(xPt, yPt, wPt, hPt, 0) & VxCell("ResizeMode", "0") & _
        "<Cell N='ImgOffsetX' V='0' F='ImgWidth*0'/><Cell N='ImgOffsetY' V='0' F='ImgHeight*0'/>" & _
        "<Cell N='ImgWidth' V='" & w & "' F='Width*1'/><Cell N='ImgHeight' V='" & h & "' F='Height*1'/>" & _
        VxCell("LinePattern", "0") & VxCell("FillPattern", "0") & _
        "<Section N='Geometry' IX='0'><Cell N='NoFill' V='0'/><Cell N='NoLine' V='1'/><Cell N='NoShow' V='0'/>" & _
        "<Row T='RelMoveTo' IX='1'><Cell N='X' V='0'/><Cell N='Y' V='0'/></Row><Row T='RelLineTo' IX='2'><Cell N='X' V='1'/><Cell N='Y' V='0'/></Row>" & _
        "<Row T='RelLineTo' IX='3'><Cell N='X' V='1'/><Cell N='Y' V='1'/></Row><Row T='RelLineTo' IX='4'><Cell N='X' V='0'/><Cell N='Y' V='1'/></Row>" & _
        "<Row T='RelLineTo' IX='5'><Cell N='X' V='0'/><Cell N='Y' V='0'/></Row></Section>" & _
        "<ForeignData ForeignType='EnhMetaFile' ObjectWidth='" & w & "' ObjectHeight='" & h & "'><Rel r:id='rId" & gVxMedia.Count & "'/></ForeignData></Shape>"
End Function

' Shape frame cells. x, y: bottom-left corner relative to the parent (points); rot: PowerPoint degrees (clockwise).
Private Function VxXform(ByVal xPt As Double, ByVal yPt As Double, ByVal wPt As Double, ByVal hPt As Double, _
                         ByVal rot As Double) As String
    Dim w As Double, h As Double
    w = wPt / 72#
    h = hPt / 72#
    If w < 0.0001 Then w = 0.0001
    If h < 0.0001 Then h = 0.0001
    VxXform = VxCell("PinX", NumXml(xPt / 72# + w / 2)) & VxCell("PinY", NumXml(yPt / 72# + h / 2)) & _
        VxCell("Width", NumXml(w)) & VxCell("Height", NumXml(h)) & _
        "<Cell N='LocPinX' V='" & NumXml(w / 2) & "' F='Width*0.5'/><Cell N='LocPinY' V='" & NumXml(h / 2) & "' F='Height*0.5'/>" & _
        VxCell("Angle", NumXml(-rot * 3.14159265358979 / 180#)) & VxCell("FlipX", "0") & VxCell("FlipY", "0")
End Function

Private Function VxCell(ByVal n As String, ByVal v As String, Optional ByVal ux As String = "") As String
    If Len(ux) > 0 Then
        VxCell = "<Cell N='" & n & "' V='" & v & "' U='" & ux & "'/>"
    Else
        VxCell = "<Cell N='" & n & "' V='" & v & "'/>"
    End If
End Function

Private Function VxName(ByVal s As Object) As String
    Dim nm As String
    On Error Resume Next
    nm = s.Name
    On Error GoTo 0
    If Len(nm) = 0 Then nm = "Shape " & (gVxId + 1)
    VxName = XmlAttr(nm)
End Function

Private Function PlaceholderIsText(ByVal s As Object) As Boolean
    Dim ct As Long
    On Error Resume Next
    ct = -1
    ct = s.PlaceholderFormat.ContainedType
    If ct = -1 Then
        PlaceholderIsText = (s.HasTextFrame = MSO_TRUE)
    Else
        PlaceholderIsText = (ct = MSO_AUTOSHAPE Or ct = MSO_TEXTBOX Or ct = MSO_PLACEHOLDER)
    End If
End Function

' A top-level copy of s on the scratch slide (for shapes inside groups or on other slides).
Private Function VxTopCopy(ByVal s As Object) As Object
    Dim r As Object
    On Error Resume Next
    s.Copy
    DoEvents
    Set r = gVxScratch.Shapes.Paste
    If r Is Nothing Then
        Err.Clear
        DoEvents
        Set r = gVxScratch.Shapes.Paste
    End If
    If Not r Is Nothing Then Set VxTopCopy = r.Item(1)
End Function

' ---- fill, line, arrowheads ---------------------------------------------------
Private Function VxStyle(ByVal s As Object, ByVal t As Long) As String
    Dim r As String, ft As Long, v As Long, c As Long
    On Error Resume Next
    v = MSO_FALSE
    If t <> MSO_LINE Then v = s.Fill.Visible
    If v = MSO_TRUE And Not gGeoOpenOnly Then
        ft = s.Fill.Type
        c = s.Fill.ForeColor.RGB
        If ft = 3 Then c = s.Fill.GradientStops.Item(1).Color.RGB
        If ft = 5 Then c = RGB(255, 255, 255)                     ' slide background fill
        r = VxCell("FillForegnd", HexRGB(c)) & VxCell("FillPattern", "1") & _
            VxCell("FillForegndTrans", NumXml(s.Fill.Transparency))
    Else
        r = VxCell("FillPattern", "0")
    End If
    v = MSO_FALSE
    v = s.Line.Visible
    If v = MSO_TRUE Then
        r = r & VxCell("LineWeight", NumXml(MaxS(s.Line.Weight, 0.1) / 72#), "PT") & _
            VxCell("LineColor", HexRGB(s.Line.ForeColor.RGB)) & _
            VxCell("LinePattern", CStr(VisioDash(s.Line.DashStyle))) & _
            VxCell("LineColorTrans", NumXml(s.Line.Transparency))
        r = r & VxArrow("Begin", s.Line.BeginArrowheadStyle, s.Line.BeginArrowheadLength, s.Line.BeginArrowheadWidth)
        r = r & VxArrow("End", s.Line.EndArrowheadStyle, s.Line.EndArrowheadLength, s.Line.EndArrowheadWidth)
    Else
        r = r & VxCell("LinePattern", "0")
    End If
    VxStyle = r
End Function

Private Function VxArrow(ByVal which As String, ByVal style As Long, ByVal ln As Long, ByVal wd As Long) As String
    Dim a As Long, sz As Long
    Select Case style
        Case 2: a = 4          ' triangle
        Case 3: a = 1          ' open
        Case 4: a = 5          ' stealth
        Case 5: a = 22         ' diamond
        Case 6: a = 10         ' oval
        Case Else: Exit Function
    End Select
    sz = ln
    If wd > sz Then sz = wd
    If sz < 1 Or sz > 3 Then sz = 2
    VxArrow = VxCell(which & "Arrow", CStr(a)) & VxCell(which & "ArrowSize", CStr(sz))
End Function

Private Function VisioDash(ByVal d As Long) As Long
    Select Case d
        Case 2, 11: VisioDash = 3         ' square dot / system dot
        Case 3: VisioDash = 10            ' round dot
        Case 4, 10: VisioDash = 2         ' dash
        Case 5, 12: VisioDash = 4         ' dash dot
        Case 6: VisioDash = 5             ' dash dot dot
        Case 7: VisioDash = 9             ' long dash
        Case 8: VisioDash = 7             ' long dash dot
        Case Else: VisioDash = 1
    End Select
End Function

' Linear gradient fill (Visio 2013+); the solid colour above remains the fallback.
Private Function VxGradient(ByVal s As Object, ByVal t As Long) As String
    Dim gs As Object, i As Long, rows As String, ang As Double
    On Error GoTo Done
    If t = MSO_LINE Or gGeoOpenOnly Then Exit Function
    If s.Fill.Visible <> MSO_TRUE Then Exit Function
    If s.Fill.Type <> 3 Then Exit Function
    Set gs = s.Fill.GradientStops
    If gs.Count < 2 Then Exit Function
    For i = 1 To gs.Count
        rows = rows & "<Row IX='" & (i - 1) & "'>" & VxCell("GradientStopColor", HexRGB(gs.Item(i).Color.RGB)) & _
               VxCell("GradientStopColorTrans", NumXml(gs.Item(i).Transparency)) & _
               VxCell("GradientStopPosition", NumXml(gs.Item(i).Position)) & "</Row>"
    Next i
    ang = 0
    ang = s.Fill.GradientAngle
    VxGradient = VxCell("FillGradientEnabled", "1") & VxCell("FillGradientDir", "0") & _
                 VxCell("FillGradientAngle", NumXml(-ang * 3.14159265358979 / 180#)) & _
                 VxCell("RotateGradientWithShape", "1") & "<Section N='FillGradient'>" & rows & "</Section>"
Done:
End Function

' ---- geometry ------------------------------------------------------------------
' Builds gGeo (Geometry sections). Coordinates passed to the Geo* helpers are PowerPoint-style:
' points, origin at the top-left of the unrotated frame, y downwards; flips are applied here.
Private Function VxGeometry(ByVal s As Object, ByVal t As Long, ByVal w As Double, ByVal h As Double) As Boolean
    Dim isConn As Boolean, fh As Boolean, fv As Boolean, ast As Long
    On Error Resume Next
    fh = (s.HorizontalFlip = MSO_TRUE)
    fv = (s.VerticalFlip = MSO_TRUE)
    isConn = (s.Connector = MSO_TRUE)
    ast = 1
    ast = s.AutoShapeType
    On Error GoTo 0
    GeoStart w, h, fh, fv
    If isConn Then
        VxGeometry = VxConnector(s, w, h)
    ElseIf t = MSO_LINE Then
        GeoBegin True, False
        GeoMove 0, 0
        GeoLine w, h
        GeoEnd False
        gGeoOpenOnly = True
        VxGeometry = True
    ElseIf t = MSO_FREEFORM Then
        If (fh Or fv Or s.Rotation <> 0) Then
            VxGeometry = VxMerged(s, True)
            If Not VxGeometry Then VxGeometry = VxNodes(s, s.Left, s.Top, True)
        Else
            VxGeometry = VxNodes(s, s.Left, s.Top, True)
        End If
    ElseIf t = MSO_TEXTBOX Or t = MSO_PLACEHOLDER Or t = MSO_TEXTEFFECT Then
        GeoRect 0, 0, w, h, 0, 0
        VxGeometry = True
    Else
        VxGeometry = VxPresetGeo(s, ast, w, h)
        If Not VxGeometry Then
            GeoStart w, h, fh, fv
            VxGeometry = VxMerged(s, False)
        End If
    End If
End Function

Private Sub GeoStart(ByVal w As Double, ByVal h As Double, ByVal fh As Boolean, ByVal fv As Boolean)
    gGeoW = w
    gGeoH = h
    gGeoFlipH = fh
    gGeoFlipV = fv
    gGeo = ""
    gGeoIx = 0
    gGeoOpenOnly = False
    gGeoSec = ""
End Sub

Private Sub GeoBegin(ByVal noFill As Boolean, ByVal noLine As Boolean)
    gGeoSec = ""
    gGeoRow = 0
    gGeoNoFill = noFill
    gGeoNoLine = noLine
    gGeoClosed = False
End Sub

Private Sub GeoXY(ByVal ux As Double, ByVal v As Double, ByRef x As Double, ByRef y As Double)
    If gGeoFlipH Then ux = gGeoW - ux
    If gGeoFlipV Then v = gGeoH - v
    x = ux / 72#
    y = (gGeoH - v) / 72#
End Sub

Private Sub GeoMove(ByVal ux As Double, ByVal v As Double)
    Dim x As Double, y As Double
    GeoXY ux, v, x, y
    gGeoRow = gGeoRow + 1
    gGeoSec = gGeoSec & "<Row T='MoveTo' IX='" & gGeoRow & "'>" & VxCell("X", NumXml(x)) & VxCell("Y", NumXml(y)) & "</Row>"
    gGeoX0 = x: gGeoY0 = y
    gGeoLX = x: gGeoLY = y
    gGeoCU = ux: gGeoCV = v: gGeoSU = ux: gGeoSV = v
End Sub

Private Sub GeoLine(ByVal ux As Double, ByVal v As Double)
    Dim x As Double, y As Double
    GeoXY ux, v, x, y
    gGeoRow = gGeoRow + 1
    gGeoSec = gGeoSec & "<Row T='LineTo' IX='" & gGeoRow & "'>" & VxCell("X", NumXml(x)) & VxCell("Y", NumXml(y)) & "</Row>"
    gGeoLX = x: gGeoLY = y
    gGeoCU = ux: gGeoCV = v
End Sub

' Closes the current sub-path (a later GeoMove starts the next one in the same section).
Private Sub GeoClose()
    If gGeoRow = 0 Then Exit Sub
    If Abs(gGeoLX - gGeoX0) > 0.000001 Or Abs(gGeoLY - gGeoY0) > 0.000001 Then
        gGeoRow = gGeoRow + 1
        gGeoSec = gGeoSec & "<Row T='LineTo' IX='" & gGeoRow & "'>" & VxCell("X", NumXml(gGeoX0)) & VxCell("Y", NumXml(gGeoY0)) & "</Row>"
        gGeoLX = gGeoX0: gGeoLY = gGeoY0
    End If
    gGeoCU = gGeoSU: gGeoCV = gGeoSV
    gGeoClosed = True
End Sub

' Quadratic Bezier from the current point, written as the equivalent cubic.
Private Sub GeoQuad(ByVal u1 As Double, ByVal v1 As Double, ByVal ux As Double, ByVal v As Double)
    GeoCubic gGeoCU + 2# / 3# * (u1 - gGeoCU), gGeoCV + 2# / 3# * (v1 - gGeoCV), _
             ux + 2# / 3# * (u1 - ux), v + 2# / 3# * (v1 - v), ux, v
End Sub

' DrawingML arcTo from the current point: radii wr / hr, start and swing angles in 1/60000 degree.
' The angles are visual angles; the point on the ellipse uses the parametric angle
' atan2(wr * sin a, hr * cos a). Written as Visio elliptical arcs of at most 90 degrees each,
' so rounded corners stay true arcs in Visio.
Private Sub GeoArcTo(ByVal wr As Double, ByVal hr As Double, ByVal stA As Double, ByVal swA As Double)
    Dim st As Double, sw As Double, t0 As Double, t1 As Double, dt As Double, cx As Double, cy As Double
    Dim n As Long, i As Long, a As Double, b As Double, m As Double
    If wr <= 0.0001 Or hr <= 0.0001 Or Abs(swA) < 1 Then Exit Sub
    st = stA * PDEG
    sw = swA * PDEG
    t0 = Atan2X(wr * Sin(st), hr * Cos(st))
    t1 = Atan2X(wr * Sin(st + sw), hr * Cos(st + sw))
    If Abs(sw) >= PI2 - 0.0000001 Then
        dt = Sgn(sw) * PI2
    Else
        dt = t1 - t0
        If sw > 0 Then
            Do While dt <= 0: dt = dt + PI2: Loop
        Else
            Do While dt >= 0: dt = dt - PI2: Loop
        End If
    End If
    cx = gGeoCU - wr * Cos(t0)
    cy = gGeoCV - hr * Sin(t0)
    n = Int(Abs(dt) / (PI2 / 4#) - 0.000000001) + 1
    For i = 1 To n
        a = t0 + dt * (i - 1) / n
        b = t0 + dt * i / n
        m = (a + b) / 2#
        GeoEArc cx + wr * Cos(m), cy + hr * Sin(m), cx + wr * Cos(b), cy + hr * Sin(b), wr, hr
    Next i
End Sub

' One Visio EllipticalArcTo row: end point, a point on the arc, axis angle and axis ratio.
Private Sub GeoEArc(ByVal mu As Double, ByVal mv As Double, ByVal eu As Double, ByVal ev As Double, _
                    ByVal rx As Double, ByVal ry As Double)
    Dim mx As Double, my As Double, x As Double, y As Double, c As String, d As Double
    GeoXY mu, mv, mx, my
    GeoXY eu, ev, x, y
    If rx >= ry Then
        c = "0": d = rx / ry
    Else
        c = NumXml(PI2 / 4#): d = ry / rx
    End If
    gGeoRow = gGeoRow + 1
    gGeoSec = gGeoSec & "<Row T='EllipticalArcTo' IX='" & gGeoRow & "'>" & VxCell("X", NumXml(x)) & VxCell("Y", NumXml(y)) & _
              VxCell("A", NumXml(mx)) & VxCell("B", NumXml(my)) & VxCell("C", c) & VxCell("D", NumXml(d)) & "</Row>"
    gGeoLX = x: gGeoLY = y
    gGeoCU = eu: gGeoCV = ev
End Sub

Private Function Atan2X(ByVal y As Double, ByVal x As Double) As Double
    If x > 0 Then
        Atan2X = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then Atan2X = Atn(y / x) + PI2 / 2# Else Atan2X = Atn(y / x) - PI2 / 2#
    ElseIf y > 0 Then
        Atan2X = PI2 / 4#
    ElseIf y < 0 Then
        Atan2X = -PI2 / 4#
    End If
End Function

Private Sub GeoCubic(ByVal u1 As Double, ByVal v1 As Double, ByVal u2 As Double, ByVal v2 As Double, _
                     ByVal ux As Double, ByVal v As Double)
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, x As Double, y As Double, f As String
    GeoXY u1, v1, x1, y1
    GeoXY u2, v2, x2, y2
    GeoXY ux, v, x, y
    f = "NURBS(1,3,1,1," & NumXml(x1) & "," & NumXml(y1) & ",0,1," & NumXml(x2) & "," & NumXml(y2) & ",0,1)"
    gGeoRow = gGeoRow + 1
    gGeoSec = gGeoSec & "<Row T='NURBSTo' IX='" & gGeoRow & "'>" & VxCell("X", NumXml(x)) & VxCell("Y", NumXml(y)) & _
              VxCell("A", "1") & VxCell("B", "1") & VxCell("C", "0") & VxCell("D", "1") & _
              "<Cell N='E' V='" & f & "' F='" & f & "'/></Row>"
    gGeoLX = x: gGeoLY = y
    gGeoCU = ux: gGeoCV = v
End Sub

Private Sub GeoEllipse(ByVal cu As Double, ByVal cv As Double, ByVal ru As Double, ByVal rv As Double)
    Dim x As Double, y As Double
    GeoBegin False, False
    GeoXY cu, cv, x, y
    gGeoRow = 1
    gGeoSec = "<Row T='Ellipse' IX='1'>" & VxCell("X", NumXml(x)) & VxCell("Y", NumXml(y)) & _
              VxCell("A", NumXml(x + ru / 72#)) & VxCell("B", NumXml(y)) & VxCell("C", NumXml(x)) & _
              VxCell("D", NumXml(y + rv / 72#)) & "</Row>"
    gGeoX0 = x: gGeoY0 = y: gGeoLX = x: gGeoLY = y
    GeoEnd True
End Sub

Private Sub GeoEnd(ByVal closed As Boolean)
    Dim nf As Long, nl As Long
    If gGeoRow = 0 Then Exit Sub
    If closed And (Abs(gGeoLX - gGeoX0) > 0.000001 Or Abs(gGeoLY - gGeoY0) > 0.000001) Then
        gGeoRow = gGeoRow + 1
        gGeoSec = gGeoSec & "<Row T='LineTo' IX='" & gGeoRow & "'>" & VxCell("X", NumXml(gGeoX0)) & VxCell("Y", NumXml(gGeoY0)) & "</Row>"
    End If
    If gGeoNoFill Or Not closed Then nf = 1
    If gGeoNoLine Then nl = 1
    gGeo = gGeo & "<Section N='Geometry' IX='" & gGeoIx & "'>" & VxCell("NoFill", CStr(nf)) & VxCell("NoLine", CStr(nl)) & _
           VxCell("NoShow", "0") & gGeoSec & "</Section>"
    gGeoIx = gGeoIx + 1
    gGeoSec = ""
    gGeoRow = 0
End Sub

' Closed polygon from a flat list u1, v1, u2, v2, ...
Private Sub GeoPoly(ParamArray c() As Variant)
    Dim i As Long
    GeoBegin False, False
    GeoMove CDbl(c(0)), CDbl(c(1))
    For i = 2 To UBound(c) - 1 Step 2
        GeoLine CDbl(c(i)), CDbl(c(i + 1))
    Next i
    GeoEnd True
End Sub

' Rectangle with optional elliptical corners (rx, ry).
Private Sub GeoRect(ByVal l As Double, ByVal t As Double, ByVal w As Double, ByVal h As Double, _
                    ByVal rx As Double, ByVal ry As Double)
    Dim kx As Double, ky As Double, r As Double, b As Double
    r = l + w
    b = t + h
    If rx > w / 2 Then rx = w / 2
    If ry > h / 2 Then ry = h / 2
    GeoBegin False, False
    If rx <= 0 Or ry <= 0 Then
        GeoMove l, t: GeoLine r, t: GeoLine r, b: GeoLine l, b
        GeoEnd True
        Exit Sub
    End If
    kx = rx * KAPPA
    ky = ry * KAPPA
    GeoMove l + rx, t
    GeoLine r - rx, t
    GeoCubic r - rx + kx, t, r, t + ry - ky, r, t + ry
    GeoLine r, b - ry
    GeoCubic r, b - ry + ky, r - rx + kx, b, r - rx, b
    GeoLine l + rx, b
    GeoCubic l + rx - kx, b, l, b - ry + ky, l, b - ry
    GeoLine l, t + ry
    GeoCubic l, t + ry - ky, l + rx - kx, t, l + rx, t
    GeoEnd True
End Sub

Private Function Adj(ByVal s As Object, ByVal i As Long, ByVal defVal As Double) As Double
    Adj = defVal
    On Error Resume Next
    If s.Adjustments.Count >= i Then Adj = s.Adjustments.Item(i)
End Function

' ---- preset geometry engine ---------------------------------------------------
' Every PowerPoint AutoShape is drawn from its DrawingML preset definition (ECMA-376
' presetShapeDefinitions, compiled by build/make_presets.py into PresetSource): the guide
' formulas are evaluated with the shape's size and adjustment values and the paths are
' written as Visio geometry - straight lines, Bezier curves and true elliptical arcs.
Private Function VxPresetGeo(ByVal s As Object, ByVal ast As Long, ByVal w As Double, ByVal h As Double) As Boolean
    Dim src As String, sec() As String, part() As String, i As Long, n As Long, d As String
    Dim v As Double, a As Double, isAng As Boolean, anyFill As Boolean, rc() As String
    src = PresetSource(ast)
    If Len(src) = 0 Then Exit Function
    sec = Split(src, "|")
    If UBound(sec) < 3 Then Exit Function
    gPw = w
    gPh = h
    ReDim gPa(0 To 0)
    If Len(sec(0)) > 0 Then
        part = Split(sec(0), ",")
        ReDim gPa(0 To UBound(part) + 1)
        n = 0
        On Error Resume Next
        n = s.Adjustments.Count
        On Error GoTo 0
        For i = 0 To UBound(part)
            d = part(i)
            isAng = (Left$(d, 1) = "!")
            If isAng Then d = Mid$(d, 2)
            v = Val(d)
            If n >= i + 1 Then
                ' PowerPoint reports linear adjustments as fractions and angles in degrees
                On Error Resume Next
                a = s.Adjustments.Item(i + 1)
                If Err.Number = 0 Then
                    If isAng Then
                        If Abs(a) > 400 Then v = a Else v = a * 60000#
                        If v < 0 And Val(d) >= 0 Then v = v + 21600000#
                    Else
                        If Abs(a) > 10 Then v = a Else v = a * 100000#
                    End If
                End If
                Err.Clear
                On Error GoTo 0
            End If
            gPa(i + 1) = v
        Next i
    End If
    If Len(sec(1)) > 0 Then
        part = Split(sec(1), ";")
        ReDim gPv(0 To UBound(part))
        For i = 0 To UBound(part)
            gPv(i) = PresetEval(part(i))
        Next i
    Else
        ReDim gPv(0 To 0)
    End If
    part = Split(sec(2), "~")
    For i = 0 To UBound(part)
        If PresetPath(part(i)) Then anyFill = True
    Next i
    If Not anyFill Then gGeoOpenOnly = True
    rc = Split(sec(3), ",")
    If Abs(PresetArg(rc(0))) > 0.5 Or Abs(PresetArg(rc(1))) > 0.5 Or Abs(PresetArg(rc(2)) - w) > 0.5 Or _
       Abs(PresetArg(rc(3)) - h) > 0.5 Then gGeoTextBox = True      ' text area is inset
    VxPresetGeo = (Len(gGeo) > 0)
End Function

' One path: "<fill f|n>,<stroke 1|0>,<path w>,<path h>:<commands>". Returns True when it is filled.
Private Function PresetPath(ByVal p As String) As Boolean
    Dim k As Long, hd() As String, cmds() As String, c() As String, i As Long, sx As Double, sy As Double
    Dim filled As Boolean
    k = InStr(p, ":")
    If k = 0 Then Exit Function
    hd = Split(Left$(p, k - 1), ",")
    cmds = Split(Mid$(p, k + 1), ";")
    filled = (hd(0) = "f")
    sx = 1: sy = 1
    If Val(hd(2)) > 0 Then sx = gPw / Val(hd(2))
    If Val(hd(3)) > 0 Then sy = gPh / Val(hd(3))
    GeoBegin Not filled, (hd(1) = "0")
    For i = 0 To UBound(cmds)
        c = Split(cmds(i), ",")
        Select Case c(0)
            Case "M": GeoMove PresetArg(c(1)) * sx, PresetArg(c(2)) * sy
            Case "L": GeoLine PresetArg(c(1)) * sx, PresetArg(c(2)) * sy
            Case "C": GeoCubic PresetArg(c(1)) * sx, PresetArg(c(2)) * sy, PresetArg(c(3)) * sx, PresetArg(c(4)) * sy, _
                               PresetArg(c(5)) * sx, PresetArg(c(6)) * sy
            Case "Q": GeoQuad PresetArg(c(1)) * sx, PresetArg(c(2)) * sy, PresetArg(c(3)) * sx, PresetArg(c(4)) * sy
            Case "A": GeoArcTo PresetArg(c(1)) * sx, PresetArg(c(2)) * sy, PresetArg(c(3)), PresetArg(c(4))
            Case "Z": GeoClose
        End Select
    Next i
    PresetPath = filled And gGeoClosed
    GeoEnd gGeoClosed
End Function

Private Function PresetArg(ByVal tok As String) As Double
    Dim c As String
    c = Left$(tok, 1)
    If c = "$" Then
        PresetArg = gPv(CLng(Mid$(tok, 2)))
    ElseIf c = "@" Then
        PresetArg = gPa(CLng(Mid$(tok, 2)))
    ElseIf c = "-" Or (c >= "0" And c <= "9") Then
        PresetArg = Val(tok)
    Else
        PresetArg = PresetBuiltin(tok)
    End If
End Function

Private Function PresetBuiltin(ByVal tok As String) As Double
    Dim ss As Double
    ss = gPw
    If gPh < ss Then ss = gPh
    Select Case tok
        Case "w", "r": PresetBuiltin = gPw
        Case "h", "b": PresetBuiltin = gPh
        Case "l", "t": PresetBuiltin = 0
        Case "hc": PresetBuiltin = gPw / 2
        Case "vc": PresetBuiltin = gPh / 2
        Case "ss": PresetBuiltin = ss
        Case "ls": If gPw > gPh Then PresetBuiltin = gPw Else PresetBuiltin = gPh
        Case "cd2": PresetBuiltin = 10800000#
        Case "cd4": PresetBuiltin = 5400000#
        Case "cd8": PresetBuiltin = 2700000#
        Case "k34": PresetBuiltin = 16200000#
        Case "k38": PresetBuiltin = 8100000#
        Case "k58": PresetBuiltin = 13500000#
        Case "k78": PresetBuiltin = 18900000#
        Case Else
            If Left$(tok, 3) = "ssd" Then
                PresetBuiltin = ss / Val(Mid$(tok, 4))
            ElseIf Left$(tok, 2) = "wd" Then
                PresetBuiltin = gPw / Val(Mid$(tok, 3))
            ElseIf Left$(tok, 2) = "hd" Then
                PresetBuiltin = gPh / Val(Mid$(tok, 3))
            End If
    End Select
End Function

' Guide formula "<op>,<x>,<y>,<z>" (DrawingML operators, angles in 1/60000 degree).
Private Function PresetEval(ByVal f As String) As Double
    Dim t() As String, x As Double, y As Double, z As Double, n As Long
    t = Split(f, ",")
    n = UBound(t)
    If n >= 1 Then x = PresetArg(t(1))
    If n >= 2 Then y = PresetArg(t(2))
    If n >= 3 Then z = PresetArg(t(3))
    Select Case t(0)
        Case "m": If z <> 0 Then PresetEval = x * y / z
        Case "a": PresetEval = x + y - z
        Case "d": If z <> 0 Then PresetEval = (x + y) / z
        Case "q": If x > 0 Then PresetEval = y Else PresetEval = z
        Case "b": PresetEval = Abs(x)
        Case "t": PresetEval = Atan2X(y, x) / PDEG
        Case "c": PresetEval = x * Cos(Atan2X(z, y))
        Case "o": PresetEval = x * Cos(y * PDEG)
        Case "x": If x > y Then PresetEval = x Else PresetEval = y
        Case "n": If x < y Then PresetEval = x Else PresetEval = y
        Case "h": PresetEval = Sqr(x * x + y * y + z * z)
        Case "p"
            If y < x Then
                PresetEval = x
            ElseIf y > z Then
                PresetEval = z
            Else
                PresetEval = y
            End If
        Case "s": PresetEval = x * Sin(Atan2X(z, y))
        Case "i": PresetEval = x * Sin(y * PDEG)
        Case "r": If x > 0 Then PresetEval = Sqr(x)
        Case "g": PresetEval = x * Tan(y * PDEG)
        Case "v": PresetEval = x
    End Select
End Function

' Connectors: straight, elbow (bentConnector2-5) and curved, from their adjustment values.
Private Function VxConnector(ByVal s As Object, ByVal w As Double, ByVal h As Double) As Boolean
    Dim ct As Long, n As Long, a1 As Double, a2 As Double, a3 As Double, x1 As Double, x2 As Double, x3 As Double
    On Error Resume Next
    ct = 1
    ct = s.ConnectorFormat.Type
    n = s.Adjustments.Count
    On Error GoTo 0
    gGeoOpenOnly = True
    GeoBegin True, False
    GeoMove 0, 0
    If ct = 2 Then
        a1 = Adj(s, 1, 0.5): a2 = Adj(s, 2, 0.5): a3 = Adj(s, 3, 0.5)
        Select Case n
            Case 0
                GeoLine w, 0: GeoLine w, h
            Case 1
                GeoLine a1 * w, 0: GeoLine a1 * w, h: GeoLine w, h
            Case 2
                GeoLine a1 * w, 0: GeoLine a1 * w, a2 * h: GeoLine w, a2 * h: GeoLine w, h
            Case Else
                GeoLine a1 * w, 0: GeoLine a1 * w, a2 * h: GeoLine a3 * w, a2 * h: GeoLine a3 * w, h: GeoLine w, h
        End Select
    ElseIf ct = 3 Then
        If n = 0 Then
            GeoCubic w / 2, 0, w, h / 2, w, h
        Else
            x2 = Adj(s, 1, 0.5) * w
            x1 = x2 / 2
            x3 = (w + x2) / 2
            GeoCubic x1, 0, x2, h / 4, x2, h / 2
            GeoCubic x2, h * 3 / 4, x3, h, w, h
        End If
    Else
        GeoLine w, h
    End If
    GeoEnd False
    VxConnector = True
End Function

' Freeform nodes -> geometry. (l, t): origin of the frame the node coordinates are measured from.
' PowerPoint stores a Bezier segment as two control nodes followed by its end vertex; the
' SegmentType of a vertex describes the segment that follows it (ShapeNodes.SetSegmentType),
' control nodes report msoSegmentCurve. Older builds were seen to flag the segment that ends at
' a node instead, so both readings are scored by tangent continuity and the better one is used.
Private Function VxNodes(ByVal s As Object, ByVal l As Double, ByVal t As Double, ByVal useFill As Boolean) As Boolean
    Dim n As Long, i As Long, p As Variant, filled As Boolean, started As Boolean, after As Boolean
    Dim px() As Double, py() As Double, sg() As Long, scoreF As Long, scoreA As Long
    Dim su As Double, sv As Double, ux As Double, v As Double
    On Error GoTo Nope
    n = s.Nodes.Count
    If n < 2 Then Exit Function
    ReDim px(1 To n): ReDim py(1 To n): ReDim sg(1 To n)
    For i = 1 To n
        p = s.Nodes.Item(i).Points
        px(i) = p(1, 1) - l
        py(i) = p(1, 2) - t
        sg(i) = s.Nodes.Item(i).SegmentType
    Next i
    filled = False
    If useFill Then filled = (s.Fill.Visible = MSO_TRUE)
    For i = 2 To n - 1
        If sg(i) = 1 And sg(i - 1) = 0 Then
            If Collinear3(px(i - 1), py(i - 1), px(i), py(i), px(i + 1), py(i + 1)) Then scoreF = scoreF + 1
            If i >= 3 Then
                If Collinear3(px(i - 2), py(i - 2), px(i - 1), py(i - 1), px(i), py(i)) Then scoreA = scoreA + 1
            End If
        End If
    Next i
    after = (scoreA > scoreF)
    i = 1
    Do While i <= n
        If Not started Then
            GeoBegin False, False
            GeoMove px(i), py(i)
            su = px(i): sv = py(i)
            ux = su: v = sv
            started = True
            i = i + 1
        Else
            If after Then
                ' flag on the node where a segment ends: node i is the first node of the next segment
                If sg(i) = 1 And i + 2 <= n Then
                    GeoCubic px(i), py(i), px(i + 1), py(i + 1), px(i + 2), py(i + 2)
                    ux = px(i + 2): v = py(i + 2)
                    i = i + 3
                Else
                    GeoLine px(i), py(i)
                    ux = px(i): v = py(i)
                    i = i + 1
                End If
            Else
                ' flag on the vertex where a segment starts (node i - 1 is the current vertex)
                If sg(i - 1) = 1 And i + 2 <= n Then
                    GeoCubic px(i), py(i), px(i + 1), py(i + 1), px(i + 2), py(i + 2)
                    ux = px(i + 2): v = py(i + 2)
                    i = i + 3
                Else
                    GeoLine px(i), py(i)
                    ux = px(i): v = py(i)
                    i = i + 1
                End If
            End If
            If i <= n And gGeoRow > 2 And Abs(ux - su) < 0.01 And Abs(v - sv) < 0.01 Then
                GeoEnd True                          ' sub-path closed: the next node starts a new one
                started = False
            End If
        End If
    Loop
    If started Then GeoEnd filled Or (Abs(ux - su) < 0.01 And Abs(v - sv) < 0.01)
    If Not filled Then gGeoOpenOnly = True
    VxNodes = (Len(gGeo) > 0)
    Exit Function
Nope:
    VxNodes = False
End Function

' True when b lies on the straight continuation from a to c (smooth join).
Private Function Collinear3(ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double, _
                            ByVal cx As Double, ByVal cy As Double) As Boolean
    Dim ux As Double, uy As Double, vx As Double, vy As Double, lu As Double, lv As Double
    ux = bx - ax: uy = by - ay
    vx = cx - bx: vy = cy - by
    lu = Sqr(ux * ux + uy * uy)
    lv = Sqr(vx * vx + vy * vy)
    If lu < 0.01 Or lv < 0.01 Then Exit Function
    Collinear3 = (Abs(ux * vy - uy * vx) / (lu * lv) < 0.02) And (ux * vx + uy * vy > 0)
End Function

' Any other closed shape: PowerPoint 2013+ "Merge Shapes - Union" of an unrotated copy with
' its duplicate gives a freeform with the shape's exact outline; its nodes are the geometry.
Private Function VxMerged(ByVal s As Object, ByVal keepOpen As Boolean) As Boolean
    Dim a As Object, b As Object, m As Object, k As Long, nmA As String, nmB As String
    Dim l As Double, t As Double, i As Long, names As String
    On Error GoTo Nope
    If gVxScratch Is Nothing Then Exit Function
    Set a = VxTopCopy(s)
    If a Is Nothing Then Exit Function
    a.Rotation = 0
    gVxCopy = gVxCopy + 1
    nmA = "vsA" & gVxCopy
    nmB = "vsB" & gVxCopy
    a.Name = nmA
    l = a.Left
    t = a.Top
    Set b = a.Duplicate.Item(1)
    b.Name = nmB
    b.Left = l
    b.Top = t
    For i = 1 To gVxScratch.Shapes.Count
        names = names & "|" & gVxScratch.Shapes.Item(i).Name & "|"
    Next i
    gVxScratch.Shapes.Range(Array(nmA, nmB)).MergeShapes 1, a      ' msoMergeUnion
    For i = gVxScratch.Shapes.Count To 1 Step -1
        If InStr(names, "|" & gVxScratch.Shapes.Item(i).Name & "|") = 0 Or gVxScratch.Shapes.Item(i).Name = nmA Then
            Set m = gVxScratch.Shapes.Item(i)
            Exit For
        End If
    Next i
    If m Is Nothing Then GoTo NopeCleanup
    gGeoFlipH = False                                ' flips are part of the merged outline
    gGeoFlipV = False
    VxMerged = VxNodes(m, l, t, Not keepOpen)
    m.Delete
    Exit Function
Nope:
    Resume NopeCleanup
NopeCleanup:
    On Error Resume Next
    If Not a Is Nothing Then a.Delete
    If Not b Is Nothing Then b.Delete
    VxMerged = False
End Function

' ---- text ----------------------------------------------------------------------
' Returns the <Text> element; cells and sects receive the text block cells and the
' Character / Paragraph sections (one row per run / paragraph, in text order).
Private Function VxText(ByVal s As Object, ByVal w As Double, ByVal h As Double, _
                        ByRef cells As String, ByRef sects As String) As String
    Dim tf As Object, tr As Object, para As Object, rn As Object
    Dim pi As Long, ri As Long, nc As Long, body As String, chars As String, paras As String
    Dim txt As String, fnt As String, fea As String, st As Long, al As Long, va As Long
    Dim sp As String, sb As Double, sa As Double, nowrap As Boolean, orient As Long, ang As Double
    Dim bx As Double, by As Double, bw As Double, bh As Double, tw As Double, th As Double, px As Double
    Dim bul As String

    cells = ""
    sects = ""
    On Error GoTo NoText
    If s.HasTextFrame <> MSO_TRUE Then Exit Function
    Set tf = s.TextFrame
    If tf.HasText <> MSO_TRUE Then Exit Function
    Set tr = tf.TextRange
    For pi = 1 To tr.Paragraphs.Count
        Set para = tr.Paragraphs(pi)
        al = 0
        Select Case para.ParagraphFormat.Alignment
            Case 2: al = 1
            Case 3: al = 2
            Case 4, 7: al = 3
            Case 5, 6: al = 4
        End Select
        sp = "-1.2"
        On Error Resume Next
        If para.ParagraphFormat.LineRuleWithin = MSO_TRUE Then
            sp = NumXml(-1.2 * para.ParagraphFormat.SpaceWithin)
        Else
            sp = NumXml(para.ParagraphFormat.SpaceWithin / 72#)
        End If
        sb = 0: sa = 0
        sb = para.ParagraphFormat.SpaceBefore
        sa = para.ParagraphFormat.SpaceAfter
        If para.ParagraphFormat.LineRuleBefore = MSO_TRUE Then sb = sb * para.Font.Size * 1.2
        If para.ParagraphFormat.LineRuleAfter = MSO_TRUE Then sa = sa * para.Font.Size * 1.2
        bul = ""
        If para.ParagraphFormat.Bullet.Visible = MSO_TRUE Then
            If para.ParagraphFormat.Bullet.Type = 2 Then
                bul = CStr(para.ParagraphFormat.Bullet.Number) & ". "
            Else
                bul = ChrW(para.ParagraphFormat.Bullet.Character) & " "
            End If
            If Len(bul) = 0 Or bul = ChrW(0) & " " Then bul = ChrW(&H2022) & " "
        End If
        On Error GoTo NoText
        If Left$(sp, 1) = "-" Then
            paras = paras & "<Row IX='" & (pi - 1) & "'>" & VxCell("HorzAlign", CStr(al)) & VxCell("SpLine", sp)
        Else
            paras = paras & "<Row IX='" & (pi - 1) & "'>" & VxCell("HorzAlign", CStr(al)) & VxCell("SpLine", sp, "PT")
        End If
        paras = paras & VxCell("SpBefore", NumXml(sb / 72#), "PT") & VxCell("SpAfter", NumXml(sa / 72#), "PT") & _
                VxCell("Bullet", "0") & "</Row>"
        body = body & "<pp IX='" & (pi - 1) & "'/>"
        For ri = 1 To para.Runs.Count
            Set rn = para.Runs(ri)
            txt = rn.Text
            Do While Len(txt) > 0 And (Right$(txt, 1) = vbCr Or Right$(txt, 1) = vbLf)
                txt = Left$(txt, Len(txt) - 1)
            Loop
            If ri = 1 Then txt = bul & txt
            If Len(txt) > 0 Then
                fnt = ThemeFontName(rn.Font.Name, False)
                fea = fnt
                On Error Resume Next
                fea = ThemeFontName(rn.Font.NameFarEast, True)
                If Len(fea) = 0 Then fea = fnt
                On Error GoTo NoText
                st = 0
                If rn.Font.Bold = MSO_TRUE Then st = st Or 1
                If rn.Font.Italic = MSO_TRUE Then st = st Or 2
                If rn.Font.Underline = MSO_TRUE Then st = st Or 4
                chars = chars & "<Row IX='" & nc & "'>" & VxCell("Font", XmlAttr(fnt)) & VxCell("AsianFont", XmlAttr(fea)) & _
                        VxCell("Color", HexRGB(rn.Font.Color.RGB)) & VxCell("Style", CStr(st)) & _
                        VxCell("Size", NumXml(rn.Font.Size / 72#), "PT") & VxCell("Pos", CStr(SuperSub(rn))) & _
                        VxCell("LangID", LangOf(txt)) & "</Row>"
                body = body & "<cp IX='" & nc & "'/>" & XmlText(Replace(txt, Chr$(11), ChrW(&H2028)))
                nc = nc + 1
            End If
        Next ri
        If pi < tr.Paragraphs.Count Then body = body & vbLf
    Next pi
    If nc = 0 Then Exit Function

    va = 1
    Select Case tf.VerticalAnchor
        Case 1, 2: va = 0
        Case 4, 5: va = 2
    End Select
    cells = VxCell("LeftMargin", NumXml(tf.MarginLeft / 72#), "PT") & VxCell("RightMargin", NumXml(tf.MarginRight / 72#), "PT") & _
            VxCell("TopMargin", NumXml(tf.MarginTop / 72#), "PT") & VxCell("BottomMargin", NumXml(tf.MarginBottom / 72#), "PT") & _
            VxCell("VerticalAlign", CStr(va))
    On Error Resume Next
    nowrap = (tf.WordWrap = MSO_FALSE)
    orient = 1
    orient = tf.Orientation
    On Error GoTo NoText
    If orient = 4 Or orient = 5 Then cells = cells & VxCell("TextDirection", "1")
    If orient = 2 Then ang = 1.5707963267949
    If orient = 3 Then ang = -1.5707963267949

    bx = 0: by = 0: bw = w: bh = h
    If gGeoTextBox And s.Rotation = 0 And ang = 0 Then
        On Error Resume Next
        bw = 0
        bx = tr.BoundLeft - s.Left
        by = tr.BoundTop - s.Top
        bw = tr.BoundWidth
        bh = tr.BoundHeight
        On Error GoTo NoText
        If bw <= 0 Or bh <= 0 Then
            bx = 0: by = 0: bw = w: bh = h
            gGeoTextBox = False
        Else
            bx = bx - bw * 0.05 - tf.MarginLeft
            bw = bw * 1.1 + tf.MarginLeft + tf.MarginRight
            by = by - tf.MarginTop
            bh = bh + tf.MarginTop + tf.MarginBottom
        End If
    End If
    If gGeoTextBox Or ang <> 0 Or nowrap Then
        tw = bw: th = bh
        If ang <> 0 Then tw = bh: th = bw
        If nowrap Then tw = MaxS(CSng(tw * 2), CSng(tw + 72))
        px = bx + bw / 2
        If nowrap And Not gGeoTextBox Then
            If al = 0 Then px = bx + tw / 2
            If al = 2 Then px = bx + bw - tw / 2
        End If
        cells = cells & VxCell("TxtPinX", NumXml(px / 72#)) & VxCell("TxtPinY", NumXml((h - by - bh / 2) / 72#)) & _
                VxCell("TxtWidth", NumXml(tw / 72#)) & VxCell("TxtHeight", NumXml(th / 72#)) & _
                VxCell("TxtLocPinX", NumXml(tw / 144#)) & VxCell("TxtLocPinY", NumXml(th / 144#)) & _
                VxCell("TxtAngle", NumXml(ang))
    End If
    sects = "<Section N='Character'>" & chars & "</Section><Section N='Paragraph'>" & paras & "</Section>"
    VxText = "<Text>" & body & "</Text>"
    Exit Function
NoText:
    cells = ""
    sects = ""
    VxText = ""
End Function

Private Function SuperSub(ByVal rn As Object) As Long
    Dim b As Double
    On Error Resume Next
    b = rn.Font.BaselineOffset
    If b > 0 Then SuperSub = 1
    If b < 0 Then SuperSub = 2
End Function

Private Function LangOf(ByVal s As String) As String
    Dim i As Long, c As Long
    LangOf = "en-US"
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        If c >= &H2E80& And c <= &H9FFF& Then LangOf = "zh-CN": Exit Function
    Next i
End Function

' Theme font placeholders (+mn-lt, +mj-ea, ...) -> the real font name of the theme.
Private Function ThemeFontName(ByVal nm As String, ByVal asian As Boolean) As String
    Dim scheme As Object, f As Object
    ThemeFontName = nm
    If Left$(nm, 1) <> "+" Then Exit Function
    On Error Resume Next
    Set scheme = gVxPres.SlideMaster.Theme.ThemeFontScheme
    If InStr(nm, "mj") > 0 Then Set f = scheme.MajorFont Else Set f = scheme.MinorFont
    If InStr(nm, "-ea") > 0 Or asian Then
        ThemeFontName = f.Item(2).Name
    Else
        ThemeFontName = f.Item(1).Name
    End If
    If Len(ThemeFontName) = 0 Or Left$(ThemeFontName, 1) = "+" Then ThemeFontName = "Calibri"
End Function

' ---- small helpers -------------------------------------------------------------
Private Sub SbReset()
    ReDim gVxSb(0 To 255)
    gVxSbN = 0
End Sub

Private Sub SbAdd(ByVal s As String)
    If gVxSbN > UBound(gVxSb) Then ReDim Preserve gVxSb(0 To UBound(gVxSb) * 2 + 1)
    gVxSb(gVxSbN) = s
    gVxSbN = gVxSbN + 1
End Sub

Private Function SbText() As String
    If gVxSbN = 0 Then Exit Function
    ReDim Preserve gVxSb(0 To gVxSbN - 1)
    SbText = Join(gVxSb, "")
End Function

Private Function HexRGB(ByVal c As Long) As String
    HexRGB = "#" & Right$("0" & Hex$(c And &HFF&), 2) & Right$("0" & Hex$((c \ &H100&) And &HFF&), 2) & _
             Right$("0" & Hex$((c \ &H10000) And &HFF&), 2)
End Function

Private Function XmlAttr(ByVal s As String) As String
    XmlAttr = Replace(Replace(XmlText(s), "'", "&apos;"), """", "&quot;")
End Function

' Escapes text for XML and drops characters XML 1.0 does not allow.
Private Function XmlText(ByVal s As String) As String
    Dim i As Long, c As Long, r As String
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        If c < 32 And c <> 9 And c <> 10 Then
            ' skip
        ElseIf c = 38 Then
            r = r & "&amp;"
        ElseIf c = 60 Then
            r = r & "&lt;"
        ElseIf c = 62 Then
            r = r & "&gt;"
        ElseIf c = &HFFFE& Or c = &HFFFF& Then
            ' skip
        Else
            r = r & Mid$(s, i, 1)
        End If
    Next i
    XmlText = r
End Function

' UTF-8 bytes of a VBA (UTF-16) string, packed into a byte string for WriteBytesX.
Private Function Utf8Bytes(ByVal s As String) As String
    Dim b() As Byte, i As Long, n As Long, c As Long, c2 As Long, k As Long
    n = Len(s)
    If n = 0 Then Exit Function
    ReDim b(0 To n * 3)
    For i = 1 To n
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        If c >= &HD800& And c <= &HDBFF& And i < n Then
            c2 = AscW(Mid$(s, i + 1, 1)) And &HFFFF&
            If c2 >= &HDC00& And c2 <= &HDFFF& Then
                c = &H10000 + (c - &HD800&) * &H400& + (c2 - &HDC00&)
                i = i + 1
            End If
        End If
        If c < &H80& Then
            b(k) = c: k = k + 1
        ElseIf c < &H800& Then
            b(k) = &HC0& Or (c \ &H40&): b(k + 1) = &H80& Or (c And &H3F&): k = k + 2
        ElseIf c < &H10000 Then
            b(k) = &HE0& Or (c \ &H1000&): b(k + 1) = &H80& Or ((c \ &H40&) And &H3F&)
            b(k + 2) = &H80& Or (c And &H3F&): k = k + 3
        Else
            If k + 4 > UBound(b) Then ReDim Preserve b(0 To UBound(b) + 16)
            b(k) = &HF0& Or (c \ &H40000): b(k + 1) = &H80& Or ((c \ &H1000&) And &H3F&)
            b(k + 2) = &H80& Or ((c \ &H40&) And &H3F&): b(k + 3) = &H80& Or (c And &H3F&): k = k + 4
        End If
    Next i
    ReDim Preserve b(0 To k - 1)
    Utf8Bytes = b
End Function

' SVG for PowerPoint versions without SVG export (Windows, PowerPoint 2013 - 2021 / LTSC):
' PowerPoint prints a copy of the drawing (alone on a temporary slide) to XPS, and the
' built-in converter turns the XPS page into SVG - paths, gradients, images and real text.
' The converter runs as a separate hidden process: the progress window shows its steps,
' Esc stops it and it is stopped after SVG_TIMEOUT seconds.
' Returns 0 on success, RUN_CANCELLED, RUN_TIMEOUT or another non-zero value on failure.
Private Function XpsToSvg(ByVal pres As Object, ByVal shp As Object, ByVal dst As String) As Long
#If Mac Then
    XpsToSvg = RUN_FAILED
#Else
    Dim tmp As Object, c As Object, rng As Object, x As String, sc As String, cmd As String, pf As String
    Dim l As Double, t As Double, w As Double, h As Double, i As Long, wasSaved As Long, st As Long
    st = RUN_FAILED
    wasSaved = pres.Saved
    x = TempFileX("xps")
    pf = TempFileX("txt")
    On Error GoTo Failed
    sc = HelperScriptPath()
    If Len(sc) = 0 Then GoTo Cleanup
    ProgressAt 0.24, Tx("step.xpsPrep")
    Set tmp = AddTempSlide(pres, shp)
    On Error Resume Next
    tmp.FollowMasterBackground = MSO_FALSE
    tmp.Background.Fill.Visible = MSO_FALSE
    tmp.DisplayMasterShapes = MSO_FALSE
    On Error GoTo Failed
    Set c = PasteFromClipboard(tmp, shp, False)
    l = 1E+30: t = 1E+30: w = -1E+30: h = -1E+30
    For i = 1 To c.Count
        If c.Item(i).Left < l Then l = c.Item(i).Left
        If c.Item(i).Top < t Then t = c.Item(i).Top
        If c.Item(i).Left + c.Item(i).Width > w Then w = c.Item(i).Left + c.Item(i).Width
        If c.Item(i).Top + c.Item(i).Height > h Then h = c.Item(i).Top + c.Item(i).Height
    Next i
    w = w - l
    h = h - t
    CheckCancel
    ProgressAt 0.3, Tx("step.xpsPrint")
    pres.PrintOptions.Ranges.ClearAll
    Set rng = pres.PrintOptions.Ranges.Add(tmp.SlideIndex, tmp.SlideIndex)
    ' Path, XPS, print intent, no frame, handout order, slides, hidden slides, range, slide range
    pres.ExportAsFixedFormat x, 1, 2, MSO_FALSE, 1, 1, MSO_TRUE, rng, 4
    If Not tmp Is Nothing Then tmp.Delete
    Set tmp = Nothing
    ProgressAt 0.4, Tx("step.xps", Tx("xps.start"))
    cmd = "& ([scriptblock]::Create([IO.File]::ReadAllText('" & PsQ(sc) & "',[Text.Encoding]::UTF8))) " & _
          "-Xps '" & PsQ(x) & "' -Out '" & PsQ(dst) & "' -Frame '" & NumXml(l) & "," & NumXml(t) & "," & _
          NumXml(w) & "," & NumXml(h) & "' -Slide '" & NumXml(pres.PageSetup.SlideWidth) & "," & _
          NumXml(pres.PageSetup.SlideHeight) & "' -Progress '" & PsQ(pf) & "' -MaxSeconds " & (SVG_TIMEOUT - 5)
    DeleteIfExists dst
    st = RunPowerShell(cmd, SVG_TIMEOUT, pf, 0.4, 0.95)
    If st = 3 Then st = RUN_TIMEOUT
    If st = 0 Then
        If Not Fso().FileExists(dst) Then st = RUN_FAILED
    End If
Cleanup:
    On Error Resume Next
    If Not tmp Is Nothing Then tmp.Delete
    pres.PrintOptions.Ranges.ClearAll
    DeleteTempX x
    DeleteTempX pf
    pres.Saved = wasSaved
    On Error GoTo 0
    XpsToSvg = st
    Exit Function
Failed:
    If Err.Number = ERR_CANCEL Then st = RUN_CANCELLED Else st = RUN_FAILED
    Resume Cleanup
#End If
End Function

#If Mac Then
#Else
' Writes the XPS -> SVG converter (embedded in this module) to the temp folder.
Private Function HelperScriptPath() As String
    Dim p As String, b() As Byte
    On Error GoTo Failed
    p = JoinPath(TempDirX(), "xps2svg-" & VS_VERSION & ".ps1")
    b = Utf8Bytes(HelperScript())
    If Fso().FileExists(p) Then
        If Fso().GetFile(p).Size <> UBound(b) + 1 Then Fso().DeleteFile p, True
    End If
    If Not Fso().FileExists(p) Then WriteBytesX p, b
    HelperScriptPath = p
    Exit Function
Failed:
    HelperScriptPath = ""
End Function
#End If


' =====================================================================
'  Temporary slide used to combine shapes and add background / margin / grid
' =====================================================================
Private Function AddTempSlide(ByVal pres As Object, ByVal src As Object) As Object
    Dim lay As Object, sld As Object, i As Long
    On Error Resume Next
    Set lay = src.Parent.CustomLayout              ' same master and theme as the source slide
    If lay Is Nothing Then Set lay = pres.SlideMaster.CustomLayouts.Item(pres.SlideMaster.CustomLayouts.Count)
    On Error GoTo 0
    Set sld = pres.Slides.AddSlide(pres.Slides.Count + 1, lay)
    For i = sld.Shapes.Count To 1 Step -1
        sld.Shapes.Item(i).Delete                  ' drop the empty layout placeholders
    Next i
    Set AddTempSlide = sld
End Function

Private Function BuildComposite(ByVal tmp As Object, ByVal src As Object, ByVal margin As Single) As Object
    Dim s As Object
    Set s = TryComposite(tmp, src, margin, False)
    If s Is Nothing Then
        ' Some objects (e.g. placeholders) cannot be grouped: use a vector EMF picture instead.
        ClearSlide tmp
        Set s = TryComposite(tmp, src, margin, True)
    End If
    If s Is Nothing Then Err.Raise vbObjectError + 513, , Tx("err.group")
    Set BuildComposite = s
End Function

Private Function TryComposite(ByVal tmp As Object, ByVal src As Object, _
                              ByVal margin As Single, ByVal asPicture As Boolean) As Object
    Dim content As Object, rect As Object, bg As Long
    PasteFromClipboard tmp, src, asPicture
    If tmp.Shapes.Count > 1 Then
        Set content = GroupAll(tmp)
        If content Is Nothing Then Exit Function
    Else
        Set content = tmp.Shapes.Item(1)
    End If

    bg = BgTypeIdx()
    If bg <> BG_NONE Or margin > 0 Then
        Set rect = tmp.Shapes.AddShape(MSO_RECTANGLE, content.Left - margin, content.Top - margin, _
                                       content.Width + 2 * margin, content.Height + 2 * margin)
        rect.Line.Visible = MSO_FALSE
        rect.Shadow.Visible = MSO_FALSE
        rect.Fill.Visible = MSO_TRUE
        rect.Fill.Solid
        If bg = BG_NONE Then
            rect.Fill.ForeColor.RGB = RGB(255, 255, 255)
            rect.Fill.Transparency = 1                  ' invisible: only reserves the margin
        Else
            rect.Fill.ForeColor.RGB = BgColorValue(bg)
            rect.Fill.Transparency = BgAlpha() / 100
        End If
        rect.ZOrder MSO_SEND_TO_BACK
        If bg = BG_GRID Then AddGrid tmp, rect
        Set content = GroupAll(tmp)
    End If
    Set TryComposite = content
End Function

' Squared-paper background: light lines every GridStep points, a darker line every 5th.
Private Sub AddGrid(ByVal tmp As Object, ByVal rect As Object)
    Dim stepPt As Single, x As Single, y As Single, i As Long, ln As Object
    Dim l As Single, t As Single, r As Single, b As Single
    l = rect.Left: t = rect.Top
    r = l + rect.Width: b = t + rect.Height
    stepPt = GridStep()
    Do While (rect.Width / stepPt + rect.Height / stepPt) > 600
        stepPt = stepPt * 2
    Loop
    i = 1
    x = l + stepPt
    Do While x < r - 0.01
        Set ln = tmp.Shapes.AddLine(x, t, x, b)
        StyleGridLine ln, (i Mod 5 = 0)
        i = i + 1
        x = l + i * stepPt
    Loop
    i = 1
    y = t + stepPt
    Do While y < b - 0.01
        Set ln = tmp.Shapes.AddLine(l, y, r, y)
        StyleGridLine ln, (i Mod 5 = 0)
        i = i + 1
        y = t + i * stepPt
    Loop
    rect.ZOrder MSO_SEND_TO_BACK
End Sub

Private Sub StyleGridLine(ByVal ln As Object, ByVal major As Boolean)
    On Error Resume Next
    ln.Shadow.Visible = MSO_FALSE
    ln.Line.Visible = MSO_TRUE
    If major Then
        ln.Line.ForeColor.RGB = RGB(176, 184, 196)
        ln.Line.Weight = 0.75
    Else
        ln.Line.ForeColor.RGB = RGB(222, 226, 232)
        ln.Line.Weight = 0.5
    End If
    ln.ZOrder MSO_SEND_TO_BACK
End Sub

' Copies src and pastes it onto sld (as shapes or as an EMF picture). Returns the pasted range.
Private Function PasteFromClipboard(ByVal sld As Object, ByVal src As Object, ByVal asPicture As Boolean) As Object
    Dim k As Long, ok As Boolean, t As Single, r As Object
    For k = 1 To 8
        src.Copy
        DoEvents
        On Error Resume Next
        Err.Clear
        If asPicture Then
            Set r = sld.Shapes.PasteSpecial(PP_PASTE_EMF)
        Else
            Set r = sld.Shapes.Paste
        End If
        ok = (Err.Number = 0)
        On Error GoTo 0
        If ok Then
            Set PasteFromClipboard = r
            Exit Function
        End If
        t = Timer
        Do While Timer - t < 0.25 And Timer >= t
            DoEvents
        Loop
    Next k
    Err.Raise vbObjectError + 512, , Tx("err.clipboard")
End Function

Private Function GroupAll(ByVal sld As Object) As Object
    On Error Resume Next
    Set GroupAll = sld.Shapes.Range().Group
    On Error GoTo 0
End Function

Private Sub ClearSlide(ByVal sld As Object)
    Dim i As Long
    For i = sld.Shapes.Count To 1 Step -1
        sld.Shapes.Item(i).Delete
    Next i
End Sub


' =====================================================================
'  Platform layer: export, files, dialogs (Windows API / macOS AppleScript)
' =====================================================================

' Shape.Export straight to the final path (macOS: via the sandbox, then moved).
Private Sub ExportRaw(ByVal shp As Object, ByVal path As String, ByVal ppFmt As Long, _
                      Optional ByVal scaleW As Long = 0, Optional ByVal scaleH As Long = 0)
#If Mac Then
    Dim t As String
    t = TempFileX(ExtOf(path))
    ExportToTemp shp, t, ppFmt, scaleW, scaleH
    MacMove t, path
#Else
    DeleteIfExists path
    ExportToTemp shp, path, ppFmt, scaleW, scaleH
#End If
End Sub

' Shape.Export to a path PowerPoint is allowed to write (on macOS: inside the sandbox).
Private Sub ExportToTemp(ByVal shp As Object, ByVal path As String, ByVal ppFmt As Long, _
                         ByVal scaleW As Long, ByVal scaleH As Long)
    On Error GoTo EH
    If scaleW > 0 And scaleH > 0 Then
        shp.Export path, ppFmt, scaleW, scaleH, PP_RELATIVE_TO_SLIDE
    Else
        shp.Export path, ppFmt
    End If
    Exit Sub
EH:
#If Mac Then
    Err.Raise Err.Number, , Err.Description & vbCrLf & Tx("err.macVersion")
#Else
    Err.Raise Err.Number, , Err.Description
#End If
End Sub

' Moves a finished temporary file to its destination (overwriting).
Private Sub MoveTempTo(ByVal t As String, ByVal dst As String)
#If Mac Then
    MacMove t, dst
#Else
    DeleteIfExists dst
    Fso().MoveFile t, dst
#End If
End Sub

Private Function TempDirX() As String
    Dim d As String
#If Mac Then
    d = Environ("HOME") & "/VecStampTemp"
    If Len(Dir(d, vbDirectory)) = 0 Then MkDir d
#Else
    d = JoinPath(Fso().GetSpecialFolder(2).Path, "VecStamp")
    If Not Fso().FolderExists(d) Then Fso().CreateFolder d
#End If
    TempDirX = d
End Function

Private Function TempFileX(ByVal ext As String) As String
    Randomize
    TempFileX = JoinPath(TempDirX(), "vs_" & Format$(Now, "hhnnss") & "_" & CStr(Int(Rnd * 1000000)) & "." & ext)
End Function

Private Sub DeleteTempX(ByVal p As String)
    On Error Resume Next
#If Mac Then
    If Len(Dir(p)) > 0 Then Kill p
#Else
    If Fso().FileExists(p) Then Fso().DeleteFile p, True
#End If
End Sub

Private Sub MakeDirX(ByVal d As String)
#If Mac Then
    If Len(Dir(d, vbDirectory)) = 0 Then MkDir d
#Else
    If Not Fso().FolderExists(d) Then Fso().CreateFolder d
#End If
End Sub

Private Sub RemoveTempDirX(ByVal d As String)
    On Error Resume Next
    If InStr(d, "VecStamp") = 0 Then Exit Sub        ' only ever delete our own temp folders
#If Mac Then
    MacScript2 "removePath", d
#Else
    Fso().DeleteFolder d, True
#End If
End Sub

' Writes ASCII text (the VSDX XML parts are pure ASCII).
Private Sub WriteTextX(ByVal p As String, ByVal txt As String)
#If Mac Then
    Dim f As Integer
    f = FreeFile
    Open p For Output As #f
    Print #f, txt;
    Close #f
#Else
    With Fso().CreateTextFile(p, True, False)
        .Write txt
        .Close
    End With
#End If
End Sub

Private Sub CopyFileX(ByVal src As String, ByVal dst As String)
#If Mac Then
    FileCopy src, dst
#Else
    Fso().CopyFile src, dst, True
#End If
End Sub

' Raw file bytes packed into a String (use LenB / MidB / InStrB on the result).
Private Function ReadBytesX(ByVal p As String) As String
    Dim b() As Byte
#If Mac Then
    Dim f As Integer, n As Long
    f = FreeFile
    Open p For Binary Access Read As #f
    n = LOF(f)
    If n > 0 Then
        ReDim b(0 To n - 1)
        Get #f, , b
    End If
    Close #f
    If n > 0 Then ReadBytesX = b
#Else
    With CreateObject("ADODB.Stream")
        .Type = 1
        .Open
        .LoadFromFile p
        If .Size > 0 Then
            b = .Read
            ReadBytesX = b
        End If
        .Close
    End With
#End If
End Function

Private Sub WriteBytesX(ByVal p As String, ByVal data As String)
    Dim b() As Byte
    b = data
#If Mac Then
    Dim f As Integer
    If Len(Dir(p)) > 0 Then Kill p
    f = FreeFile
    Open p For Binary Access Write As #f
    Put #f, , b
    Close #f
#Else
    With CreateObject("ADODB.Stream")
        .Type = 1
        .Open
        .Write b
        .SaveToFile p, 2
        .Close
    End With
#End If
End Sub

' ASCII text as a byte string, for comparisons with InStrB.
Private Function AsciiB(ByVal s As String) As String
    Dim i As Long, r As String
    For i = 1 To Len(s)
        r = r & ChrB(Asc(Mid$(s, i, 1)))
    Next i
    AsciiB = r
End Function

Private Function FileExistsX(ByVal p As String) As Boolean
#If Mac Then
    FileExistsX = (MacScript2("pathExists", p) = "1")
#Else
    FileExistsX = Fso().FileExists(p)
#End If
End Function

Private Function FolderExistsX(ByVal p As String) As Boolean
#If Mac Then
    FolderExistsX = (MacScript2("isFolder", p) = "1")
#Else
    FolderExistsX = Fso().FolderExists(p)
#End If
End Function

Private Function DesktopPathX() As String
    Dim d As String
#If Mac Then
    d = MacScript2("desktopFolder", "")
    If Len(d) > 1 And Right$(d, 1) = "/" Then d = Left$(d, Len(d) - 1)
#Else
    On Error Resume Next
    d = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    On Error GoTo 0
    If Len(d) = 0 Then d = Environ$("USERPROFILE") & "\Desktop"
#End If
    DesktopPathX = d
End Function

Private Function PickFolderX(ByVal initDir As String, ByVal promptText As String) As String
    Dim f As String
#If Mac Then
    f = MacScript2("chooseFolder", promptText & MAC_SEP & initDir)
#Else
    With Application.FileDialog(MSO_FOLDER_PICKER)
        .Title = promptText
        .AllowMultiSelect = False
        If Len(initDir) > 0 Then .InitialFileName = JoinPath(initDir, "")
        If .Show = -1 Then f = .SelectedItems(1)
    End With
#End If
    If Len(f) > 3 And Right$(f, 1) = PathSep() Then f = Left$(f, Len(f) - 1)
    PickFolderX = f
End Function

Private Sub OpenPathX(ByVal p As String)
#If Mac Then
    MacScript2 "openPath", p
#Else
    CreateObject("WScript.Shell").Run "explorer.exe """ & p & """", 1, False
#End If
End Sub

Private Sub RevealFileX(ByVal p As String)
#If Mac Then
    MacScript2 "revealFile", p
#Else
    CreateObject("WScript.Shell").Run "explorer.exe /select,""" & p & """", 1, False
#End If
End Sub

Private Function FileManagerName() As String
#If Mac Then
    FileManagerName = Tx("fm.mac")
#Else
    FileManagerName = Tx("fm.win")
#End If
End Function

' Save dialog. fmt is updated when the user picks another type / types another extension.
Private Function AskSavePath(ByVal initDir As String, ByVal baseName As String, ByRef fmt As Long) As String
    Dim p As String, k As Long
#If Mac Then
    p = MacScript2("chooseSaveFile", Tx("dlg.saveAs", UCase$(FmtExt(fmt))) & MAC_SEP & _
                   baseName & "." & FmtExt(fmt) & MAC_SEP & initDir)
    If Len(p) = 0 Then Exit Function
#Else
    Dim ofn As OPENFILENAMEW
    Dim fileBuf As String, filterStr As String, dlgTitle As String, defExt As String
    Dim r As Long, i As Long, folder As String, nm As String

    For i = 0 To FMT_LAST
        filterStr = filterStr & FilterLabel(i) & vbNullChar & "*." & FmtExt(i) & vbNullChar
    Next i
    filterStr = filterStr & vbNullChar
    fileBuf = baseName & String$(2048, vbNullChar)
    dlgTitle = AppTitle() & " - " & Tx("dlg.saveTitle")
    defExt = FmtExt(fmt)

    With ofn
        .lStructSize = LenB(ofn)
        .hwndOwner = GetActiveWindow()
        .lpstrFilter = StrPtr(filterStr)
        .nFilterIndex = fmt + 1
        .lpstrFile = StrPtr(fileBuf)
        .nMaxFile = Len(fileBuf)
        .lpstrInitialDir = StrPtr(initDir)
        .lpstrTitle = StrPtr(dlgTitle)
        .lpstrDefExt = StrPtr(defExt)
        .Flags = OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR Or OFN_PATHMUSTEXIST Or OFN_EXPLORER
    End With

    On Error GoTo Fallback
    r = GetSaveFileNameW(ofn)
    On Error GoTo 0
    If r = 0 Then Exit Function                     ' cancelled
    k = InStr(fileBuf, vbNullChar)
    If k > 0 Then p = Left$(fileBuf, k - 1) Else p = fileBuf
    k = FmtFromExt(LCase$(ExtOf(p)))
    If k < 0 Then
        fmt = ofn.nFilterIndex - 1
        If fmt < 0 Or fmt > FMT_LAST Then fmt = FMT_EMF
    End If
#End If
    k = FmtFromExt(LCase$(ExtOf(p)))
    If k >= 0 Then
        fmt = k                                     ' the extension decides the format
    Else
        p = p & "." & FmtExt(fmt)
    End If
    AskSavePath = p
    Exit Function

#If Mac Then
#Else
Fallback:
    ' Save dialog API unavailable: folder picker + file name prompt.
    folder = PickFolderX(initDir, Tx("dlg.pickFolder"))
    If Len(folder) = 0 Then Exit Function
    nm = InputBox(Tx("dlg.fileName"), AppTitle(), baseName)
    If Len(nm) = 0 Then Exit Function
    AskSavePath = JoinPath(folder, SafeName(nm) & "." & FmtExt(fmt))
#End If
End Function

Private Function GzipX(ByVal src As String, ByVal dst As String) As Boolean
    Dim ok As Boolean
#If Mac Then
    ok = (MacScript2("gzipFile", src & MAC_SEP & dst) = "ok")
#Else
    Dim cmd As String
    DeleteIfExists dst
    cmd = "$i=[IO.File]::OpenRead('" & PsQ(src) & "');" & _
          "$o=[IO.File]::Create('" & PsQ(dst) & "');" & _
          "$g=New-Object IO.Compression.GZipStream($o,[IO.Compression.CompressionMode]::Compress);" & _
          "$i.CopyTo($g);$g.Dispose();$o.Dispose();$i.Dispose()"
    ok = (RunPowerShell(cmd, 120) = 0)
    DeleteIfExists src
#End If
    If ok Then ok = FileExistsX(dst)
    GzipX = ok
End Function

' Zips a folder into an OPC package; [Content_Types].xml goes first, entry names use "/".
Private Sub ZipFolderX(ByVal dirPath As String, ByVal zipPath As String)
#If Mac Then
    If MacScript2("zipFolder", dirPath & MAC_SEP & zipPath) <> "ok" Then
        Err.Raise vbObjectError + 516, , Tx("err.zip")
    End If
#Else
    Dim cmd As String
    ProgressAt 0.94, Tx("step.zip")
    DeleteIfExists zipPath
    cmd = "Add-Type -AssemblyName System.IO.Compression;Add-Type -AssemblyName System.IO.Compression.FileSystem;" & _
          "$src='" & PsQ(dirPath) & "';$dst='" & PsQ(zipPath) & "';" & _
          "$z=[IO.Compression.ZipFile]::Open($dst,'Create');" & _
          "$fs=@(Get-ChildItem -LiteralPath $src -Recurse -File -Force|Sort-Object @{e={if($_.Name -eq '[Content_Types].xml'){0}else{1}}},FullName);" & _
          "foreach($f in $fs){$n=$f.FullName.Substring($src.Length).TrimStart('\').Replace('\','/');" & _
          "[void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($z,$f.FullName,$n)};$z.Dispose()"
    Select Case RunPowerShell(cmd, 120)
        Case 0
            If Not Fso().FileExists(zipPath) Then Err.Raise vbObjectError + 516, , Tx("err.zipPs")
        Case RUN_CANCELLED
            Err.Raise ERR_CANCEL, , "cancelled"
        Case Else
            Err.Raise vbObjectError + 516, , Tx("err.zipPs")
    End Select
#End If
End Sub

' EMF -> SVG with Inkscape or LibreOffice, when one of them is installed.
' Returns 0 on success, RUN_CANCELLED, or another non-zero value.
Private Function ConvertEmfToSvg(ByVal src As String, ByVal dst As String) As Long
    ConvertEmfToSvg = RUN_FAILED
#If Mac Then
    If MacScript2("inkscapeSvg", src & MAC_SEP & dst) = "ok" Then
        ConvertEmfToSvg = 0
    ElseIf MacScript2("libreofficeSvg", src & MAC_SEP & dst) = "ok" Then
        ConvertEmfToSvg = 0
    End If
#Else
    Dim c As Variant, exe As String, outDir As String, produced As String, st As Long
    For Each c In Array(Environ$("ProgramFiles") & "\Inkscape\bin\inkscape.exe", _
                        Environ$("ProgramW6432") & "\Inkscape\bin\inkscape.exe", _
                        Environ$("ProgramFiles(x86)") & "\Inkscape\bin\inkscape.exe", _
                        Environ$("LOCALAPPDATA") & "\Programs\Inkscape\bin\inkscape.exe")
        If Fso().FileExists(CStr(c)) Then exe = CStr(c): Exit For
    Next
    If Len(exe) > 0 Then
        ProgressAt 0.55, Tx("step.inkscape")
        st = RunHidden("""" & exe & """ """ & src & """ --export-type=svg ""--export-filename=" & dst & """", 120)
        If st = RUN_CANCELLED Then ConvertEmfToSvg = st: Exit Function
        If Fso().FileExists(dst) Then ConvertEmfToSvg = 0: Exit Function
    End If

    exe = ""
    For Each c In Array(Environ$("ProgramFiles") & "\LibreOffice\program\soffice.exe", _
                        Environ$("ProgramW6432") & "\LibreOffice\program\soffice.exe", _
                        Environ$("ProgramFiles(x86)") & "\LibreOffice\program\soffice.exe")
        If Fso().FileExists(CStr(c)) Then exe = CStr(c): Exit For
    Next
    If Len(exe) > 0 Then
        outDir = JoinPath(TempDirX(), "lo_" & Format$(Now, "hhnnss"))
        MakeDirX outDir
        ProgressAt 0.6, Tx("step.libreoffice")
        st = RunHidden("""" & exe & """ -env:UserInstallation=" & FileUrlX(JoinPath(TempDirX(), "lo_profile")) & _
                " --headless --norestore --convert-to svg --outdir """ & outDir & """ """ & src & """", 180)
        produced = JoinPath(outDir, BaseNameOf(src) & ".svg")
        If Fso().FileExists(produced) Then
            DeleteIfExists dst
            Fso().MoveFile produced, dst
            ConvertEmfToSvg = 0
        ElseIf st = RUN_CANCELLED Then
            ConvertEmfToSvg = st
        End If
        RemoveTempDirX outDir
    End If
#End If
End Function

#If Mac Then
' Calls a handler of the AppleScript helper (outside the PowerPoint sandbox).
Private Function MacScript2(ByVal handler As String, ByVal param As String) As String
    On Error GoTo EH
    MacScript2 = AppleScriptTask(MAC_SCRIPT, handler, param)
    Exit Function
EH:
    Err.Raise vbObjectError + 520, , Tx("err.macHelper", MAC_SCRIPT)
End Function

Private Sub MacMove(ByVal src As String, ByVal dst As String)
    If MacScript2("moveFile", src & MAC_SEP & dst) <> "ok" Then
        Err.Raise vbObjectError + 521, , Tx("err.macWrite", dst)
    End If
End Sub
#Else
Private Function Fso() As Object
    Static f As Object
    If f Is Nothing Then Set f = CreateObject("Scripting.FileSystemObject")
    Set Fso = f
End Function

Private Sub DeleteIfExists(ByVal p As String)
    If Fso().FileExists(p) Then Fso().DeleteFile p, True
End Sub

' Runs a program without a window and without freezing PowerPoint: while it works, the
' progress window keeps updating (from progressFile, when given), Esc ends the program and
' so does the time limit. Returns the program's exit code, RUN_FAILED (could not start),
' RUN_CANCELLED or RUN_TIMEOUT.
Private Function RunHidden(ByVal cmdLine As String, ByVal timeoutSec As Double, _
                           Optional ByVal progressFile As String = "", Optional ByVal frac0 As Double = 0, _
                           Optional ByVal frac1 As Double = 0) As Long
    Dim si As STARTUPINFOW, pi As PROCESS_INFORMATION, buf As String, code As Long, t0 As Double, el As Double
    Dim lastSec As Long
    buf = cmdLine & vbNullChar
    si.cb = LenB(si)
    si.dwFlags = STARTF_USESHOWWINDOW
    si.wShowWindow = SW_HIDE
    If CreateProcessW(0, StrPtr(buf), 0, 0, 0, CREATE_NO_WINDOW, 0, 0, si, pi) = 0 Then
        RunHidden = RUN_FAILED
        Exit Function
    End If
    t0 = Timer
    lastSec = -1
    Do
        If WaitForSingleObject(pi.hProcess, 120) <> WAIT_TIMEOUT Then Exit Do
        el = Timer - t0
        If el < 0 Then el = el + 86400#
        If Len(progressFile) > 0 Then ProgressFromFile progressFile, frac0, frac1, el
        If Int(el) <> lastSec Then
            lastSec = Int(el)
            ProgressElapsed lastSec
        End If
        DoEvents
        If ProgressCancelled() Then
            TerminateProcess pi.hProcess, 1
            WaitForSingleObject pi.hProcess, 3000
            code = RUN_CANCELLED
            GoTo Done
        End If
        If el > timeoutSec Then
            TerminateProcess pi.hProcess, 1
            WaitForSingleObject pi.hProcess, 3000
            code = RUN_TIMEOUT
            GoTo Done
        End If
    Loop
    If GetExitCodeProcess(pi.hProcess, code) = 0 Then code = RUN_FAILED
Done:
    CloseHandle pi.hThread
    CloseHandle pi.hProcess
    RunHidden = code
End Function

' Runs a PowerShell snippet (no window). The script is passed as plain text with -Command - no
' encoded command and no execution-policy switch, so security software sees ordinary use.
Private Function RunPowerShell(ByVal script As String, ByVal timeoutSec As Double, _
                               Optional ByVal progressFile As String = "", Optional ByVal frac0 As Double = 0, _
                               Optional ByVal frac1 As Double = 0) As Long
    Dim exe As String
    exe = Environ$("SystemRoot") & "\System32\WindowsPowerShell\v1.0\powershell.exe"
    If Not Fso().FileExists(exe) Then exe = "powershell.exe"
    RunPowerShell = RunHidden("""" & exe & """ -NoLogo -NoProfile -NonInteractive -Command """ & _
                              "$ErrorActionPreference='Stop';" & Replace(script, """", "'") & """", _
                              timeoutSec, progressFile, frac0, frac1)
End Function

' Shows the step the XPS -> SVG converter reports in its progress file ("percent|step|a|b").
Private Sub ProgressFromFile(ByVal path As String, ByVal frac0 As Double, ByVal frac1 As Double, ByVal el As Double)
    Dim txt As String, f() As String, pct As Double, d As String
    On Error Resume Next
    If Not Fso().FileExists(path) Then Exit Sub
    With Fso().OpenTextFile(path, 1)
        txt = .ReadAll
        .Close
    End With
    If Err.Number <> 0 Or Len(txt) = 0 Or txt = gPgFile Then Exit Sub
    On Error GoTo 0
    gPgFile = txt
    f = Split(txt, "|")
    If UBound(f) < 3 Then Exit Sub
    pct = Val(f(0)) / 100#
    Select Case f(1)
        Case "open", "write", "done": d = Tx("xps." & f(1))
        Case "page": d = Tx("xps.page", f(3))
        Case "elem": d = Tx("xps.elem", f(2), f(3))
        Case "font": d = Tx("xps.font", f(2))
        Case Else: d = f(1)
    End Select
    ProgressAt frac0 + (frac1 - frac0) * pct, Tx("step.xps", d)
End Sub

Private Function PsQ(ByVal s As String) As String
    ' Escape for a PowerShell single-quoted string (PowerShell also treats curly quotes as quotes).
    s = Replace(s, "'", "''")
    s = Replace(s, ChrW(&H2018), ChrW(&H2018) & ChrW(&H2018))
    s = Replace(s, ChrW(&H2019), ChrW(&H2019) & ChrW(&H2019))
    PsQ = s
End Function

' file:/// URL with UTF-8 percent-encoding (for LibreOffice's -env:UserInstallation).
Private Function FileUrlX(ByVal p As String) As String
    Dim b() As Byte, i As Long, c As Long, s As String
    With CreateObject("ADODB.Stream")
        .Type = 2
        .Charset = "utf-8"
        .Open
        .WriteText Replace(p, "\", "/")
        .Position = 0
        .Type = 1
        .Position = 3                               ' skip the UTF-8 byte order mark
        b = .Read
        .Close
    End With
    For i = LBound(b) To UBound(b)
        c = b(i)
        If (c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or _
           c = 45 Or c = 46 Or c = 95 Or c = 126 Or c = 47 Or c = 58 Then
            s = s & Chr$(c)
        Else
            s = s & "%" & Right$("0" & Hex$(c), 2)
        End If
    Next i
    FileUrlX = "file:///" & s
End Function
#End If


' =====================================================================
'  Progress window (Windows: a small native window; macOS: a dialog from the helper)
'  A run has items (shapes, files or slides); ProgressAt sets the position inside the current
'  item and the step being done, and the window shows the overall percentage.
' =====================================================================
Private Sub ProgressBegin(ByVal title As String, ByVal items As Long)
    ProgressEnd
    gPgItems = items
    If gPgItems < 1 Then gPgItems = 1
    gPgItem = 1
    gPgFrac = 0
    gPgActive = True
    gPgCancel = False
    gPgTitle = title
    gPgLabel = ""
    gPgDetail = Tx("prog.preparing")
    gPgFile = ""
#If Mac Then
    gPgMacPid = ""
    gPgMacPct = -100
    gPgMacTime = 0
    ProgressShow
#Else
    GetAsyncKeyState VK_ESCAPE                     ' forget an Esc pressed before the run started
    ProgressCreate title
#End If
End Sub

' Starts item i of the run, with a short label such as "2 / 5: Rectangle 3".
Private Sub ProgressItem(ByVal i As Long, ByVal label As String)
    If Not gPgActive Then Exit Sub
    gPgItem = i
    If gPgItem > gPgItems Then gPgItems = gPgItem
    gPgFrac = 0
    gPgLabel = label
    ProgressShow
End Sub

' Position inside the current item (0..1, never backwards) and the step being done.
Private Sub ProgressAt(ByVal frac As Double, ByVal detail As String)
    If Not gPgActive Then Exit Sub
    If frac < gPgFrac Then frac = gPgFrac
    If frac > 1 Then frac = 1
    gPgFrac = frac
    If Len(detail) > 0 Then gPgDetail = detail
    ProgressShow
End Sub

' While an external program runs: show how long it has been working.
Private Sub ProgressElapsed(ByVal sec As Long)
    If Not gPgActive Or sec < 2 Then Exit Sub
#If Mac Then
#Else
    ProgressShow sec
#End If
End Sub

Private Function ProgressPercent() As Long
    Dim p As Long
    p = Int(((gPgItem - 1) + gPgFrac) / gPgItems * 100# + 0.5)
    If p < 0 Then p = 0
    If p > 100 Then p = 100
    ProgressPercent = p
End Function

Private Sub ProgressEnd()
    gPgActive = False
#If Mac Then
    If Len(gPgMacPid) > 0 Then
        On Error Resume Next
        AppleScriptTask MAC_SCRIPT, "progressHide", gPgMacPid
        gPgMacPid = ""
    End If
#Else
    On Error Resume Next
    If gPgWnd <> 0 Then DestroyWindow gPgWnd
    gPgWnd = 0
    If gPgFont <> 0 Then DeleteObject gPgFont
    If gPgFontB <> 0 Then DeleteObject gPgFontB
    gPgFont = 0
    gPgFontB = 0
#End If
End Sub

' True once the user pressed Esc (Windows) or closed the progress dialog with its button
' (macOS) during the run; it stays True until the run ends.
Private Function ProgressCancelled() As Boolean
#If Mac Then
    If gPgActive And Not gPgCancel And Len(gPgMacPid) > 0 Then
        If Abs(Timer - gPgMacCheck) >= 1.5 Then
            gPgMacCheck = Timer
            On Error Resume Next
            If AppleScriptTask(MAC_SCRIPT, "progressAlive", gPgMacPid) = "0" Then gPgCancel = True
        End If
    End If
#Else
    If gPgActive And Not gPgCancel Then
        If (GetAsyncKeyState(VK_ESCAPE) And &H8001) <> 0 Then gPgCancel = True
    End If
#End If
    ProgressCancelled = gPgCancel
End Function

' Stops the run (error ERR_CANCEL, handled by the export flows) when Esc was pressed.
Private Sub CheckCancel()
    If ProgressCancelled() Then Err.Raise ERR_CANCEL, , "cancelled"
End Sub

' MsgBox that hides the progress window while it is shown.
Private Function AskUser(ByVal prompt As String, ByVal buttons As VbMsgBoxStyle) As VbMsgBoxResult
#If Mac Then
    AskUser = MsgBox(prompt, buttons, AppTitle())
#Else
    If gPgWnd <> 0 Then ShowWindow gPgWnd, SW_HIDE
    AskUser = MsgBox(prompt, buttons, AppTitle())
    If gPgWnd <> 0 And gPgActive Then ShowWindow gPgWnd, SW_SHOWNOACTIVATE
    GetAsyncKeyState VK_ESCAPE                     ' an Esc that closed the message box does not cancel
#End If
End Function

#If Mac Then
' macOS: the helper shows a small dialog; it is replaced when the text changes noticeably
' (at most every 1.5 seconds, the dialog cannot be updated in place).
Private Sub ProgressShow(Optional ByVal sec As Long = -1)
    Dim pct As Long, txt As String, now As Double
    If Not gPgActive Then Exit Sub
    pct = ProgressPercent()
    txt = gPgTitle
    If Len(gPgLabel) > 0 Then txt = txt & " - " & gPgLabel
    txt = txt & vbLf & gPgDetail & vbLf & pct & "%"
    now = Timer
    If Len(gPgMacPid) > 0 Then
        If Abs(pct - gPgMacPct) < 10 And Abs(now - gPgMacTime) < 1.5 Then Exit Sub
        If txt = gPgMacText Then Exit Sub
    End If
    On Error Resume Next
    If Len(gPgMacPid) > 0 Then
        gPgMacPid = AppleScriptTask(MAC_SCRIPT, "progressUpdate", gPgMacPid & MAC_SEP & txt & MAC_SEP & Tx("btn.cancel"))
    Else
        gPgMacPid = AppleScriptTask(MAC_SCRIPT, "progressShow", txt & MAC_SEP & Tx("btn.cancel"))
    End If
    gPgMacPct = pct
    gPgMacTime = now
    gPgMacText = txt
    gPgMacCheck = now
End Sub
#Else
Private Sub ProgressCreate(ByVal title As String)
    Dim r As RECT, dpi As Long, k As Double, w As Long, h As Long, x As Long, y As Long
    Dim cls As String, cap As String, face As String, icc As INITCOMMONCONTROLSEX_T
    Dim owner As LongPtr, hdc As LongPtr
    On Error GoTo Done
    icc.dwSize = LenB(icc)
    icc.dwICC = ICC_PROGRESS_CLASS
    InitCommonControlsEx icc
    owner = GetActiveWindow()
    hdc = GetDC(0)
    dpi = GetDeviceCaps(hdc, LOGPIXELSY)
    ReleaseDC 0, hdc
    If dpi < 72 Then dpi = 96
    k = dpi / 96#
    w = CLng(480 * k)
    h = CLng(158 * k)
    If owner <> 0 And GetWindowRect(owner, r) <> 0 Then
        x = r.Left + (r.Right - r.Left - w) \ 2
        y = r.Top + (r.Bottom - r.Top - h) \ 2
    Else
        x = (GetSystemMetrics(0) - w) \ 2
        y = (GetSystemMetrics(1) - h) \ 2
    End If
    cls = "#32770"
    cap = AppTitle()
    gPgWnd = CreateWindowExW(WS_EX_DLGMODALFRAME Or WS_EX_TOOLWINDOW, StrPtr(cls), StrPtr(cap), WS_POPUP Or WS_CAPTION, _
                             x, y, w, h, owner, 0, 0, 0)
    If gPgWnd = 0 Then Exit Sub
    face = UiFontName()
    gPgFontB = CreateFontW(-CLng(10.5 * dpi / 72), 0, 0, 0, 700, 0, 0, 0, 1, 0, 0, 5, 0, StrPtr(face))
    gPgFont = CreateFontW(-CLng(9 * dpi / 72), 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 5, 0, StrPtr(face))
    cls = "Static"
    gPgMsg = CreateWindowExW(0, StrPtr(cls), StrPtr(title), WS_CHILD Or WS_VISIBLE Or SS_NOPREFIX, _
                             CLng(20 * k), CLng(14 * k), CLng(360 * k), CLng(24 * k), gPgWnd, 0, 0, 0)
    cap = "0%"
    gPgPct = CreateWindowExW(0, StrPtr(cls), StrPtr(cap), WS_CHILD Or WS_VISIBLE Or SS_NOPREFIX Or SS_RIGHT, _
                             CLng(380 * k), CLng(14 * k), CLng(72 * k), CLng(24 * k), gPgWnd, 0, 0, 0)
    cls = "msctls_progress32"
    gPgBar = CreateWindowExW(0, StrPtr(cls), 0, WS_CHILD Or WS_VISIBLE, _
                             CLng(20 * k), CLng(44 * k), CLng(432 * k), CLng(16 * k), gPgWnd, 0, 0, 0)
    cls = "Static"
    cap = gPgDetail
    gPgSub = CreateWindowExW(0, StrPtr(cls), StrPtr(cap), WS_CHILD Or WS_VISIBLE Or SS_NOPREFIX, _
                             CLng(20 * k), CLng(68 * k), CLng(432 * k), CLng(42 * k), gPgWnd, 0, 0, 0)
    SendMessageW gPgMsg, WM_SETFONT, gPgFontB, 1
    SendMessageW gPgPct, WM_SETFONT, gPgFontB, 1
    SendMessageW gPgSub, WM_SETFONT, gPgFont, 1
    If gPgBar <> 0 Then SendMessageW gPgBar, PBM_SETRANGE32, 0, 100
    ShowWindow gPgWnd, SW_SHOWNOACTIVATE
    UpdateWindow gPgWnd
    DoEvents
Done:
End Sub

Private Sub ProgressShow(Optional ByVal sec As Long = -1)
    Dim head As String, detail As String, pct As Long
    On Error Resume Next
    If gPgWnd = 0 Then Exit Sub
    If IsWindow(gPgWnd) = 0 Then gPgWnd = 0: Exit Sub
    pct = ProgressPercent()
    head = gPgTitle
    If Len(gPgLabel) > 0 Then head = head & " - " & gPgLabel
    detail = gPgDetail
    If sec >= 2 Then detail = detail & Tx("prog.elapsed", sec)
    detail = detail & vbCrLf & Tx("prog.escHint")
    SetWindowTextW gPgMsg, StrPtr(head)
    SetWindowTextW gPgPct, StrPtr(CStr(pct) & "%")
    SetWindowTextW gPgSub, StrPtr(detail)
    If gPgBar <> 0 Then SendMessageW gPgBar, PBM_SETPOS, pct, 0
    UpdateWindow gPgWnd
    DoEvents
End Sub
#End If


' Font for native windows in the current interface language.
Private Function UiFontName() As String
    Select Case CurLang()
        Case "zh-TW": UiFontName = "Microsoft JhengHei UI"
        Case "ja": UiFontName = "Yu Gothic UI"
        Case "en": UiFontName = "Segoe UI"
        Case Else: UiFontName = "Microsoft YaHei UI"
    End Select
End Function

' A name shortened for one-line labels.
Private Function ShortName(ByVal nm As String) As String
    If Len(nm) > 28 Then nm = Left$(nm, 26) & U("2026")
    ShortName = nm
End Function


' =====================================================================
'  Feedback, About and donations
' =====================================================================
Private Sub DoFeedback()
    Dim ans As VbMsgBoxResult
    ans = MsgBox(Tx("fb.ask", VS_EMAIL), vbQuestion + vbYesNoCancel, AppTitle() & " - " & Tx("fb.title"))
    Select Case ans
        Case vbYes: FeedbackMail
        Case vbNo: FeedbackIssue
    End Select
End Sub

Private Sub FeedbackMail()
    Dim copied As Boolean, url As String
    copied = SetClipboardText(VS_EMAIL)
    If MsgBox(Tx(IIf(copied, "fb.mailCopied", "fb.mailPlain"), VS_EMAIL), _
              vbQuestion + vbYesNo + vbDefaultButton2, AppTitle() & " - " & Tx("fb.mailTitle")) <> vbYes Then Exit Sub
    url = "mailto:" & VS_EMAIL & "?subject=" & UrlEncode(Tx("fb.subject", VS_VERSION)) & _
          "&body=" & UrlEncode(vbCrLf & vbCrLf & "----" & vbCrLf & EnvironmentInfo())
    If Not OpenUrlX(url) Then
        MsgBox Tx("fb.noMailApp", VS_EMAIL), vbInformation, AppTitle()
    End If
End Sub

Private Sub FeedbackIssue()
    Dim url As String, body As String
    body = Replace(Tx("fb.issueBody", EnvironmentInfo()), vbCrLf, vbLf)
    url = VS_REPO & "/issues/new?title=" & UrlEncode(Tx("fb.issueTitle")) & "&body=" & UrlEncode(body)
    If Not OpenUrlX(url) Then
        MsgBox Tx("fb.noBrowser", VS_REPO & "/issues"), vbInformation, AppTitle()
    End If
End Sub

Private Function EnvironmentInfo() As String
    Dim s As String
    s = "VecStamp: " & VS_VERSION
    On Error Resume Next
    s = s & vbCrLf & "PowerPoint: " & Application.Version & " (build " & Application.Build & ")"
    s = s & vbCrLf & "OS: " & Application.OperatingSystem
    EnvironmentInfo = s
End Function

Private Sub DoAbout()
    MsgBox Tx("about.text", VS_VERSION, VS_AUTHOR, VS_EMAIL, VS_REPO, VS_YEAR) & vbCrLf & vbCrLf & _
           Tx("about.lang", LangName(CurLang())), vbInformation, Tx("about.title", AppTitle())
End Sub

Private Sub DoDonateThanks(ByVal which As String)
    MsgBox Tx("don.thanks", which), vbInformation, AppTitle() & " - " & Tx("don.title")
End Sub

Private Function OpenUrlX(ByVal url As String) As Boolean
    On Error GoTo Failed
#If Mac Then
    MacScript2 "openPath", url
#Else
    CreateObject("Shell.Application").ShellExecute url
#End If
    OpenUrlX = True
    Exit Function
Failed:
    OpenUrlX = False
End Function

Private Function SetClipboardText(ByVal s As String) As Boolean
    On Error GoTo Failed
#If Mac Then
    SetClipboardText = (MacScript2("setClipboard", s) = "ok")
#Else
    With CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")    ' MSForms.DataObject
        .SetText s
        .PutInClipboard
    End With
    SetClipboardText = True
#End If
    Exit Function
Failed:
    SetClipboardText = False
End Function

' Percent-encoding of the UTF-8 bytes (for mailto: and https: query strings).
Private Function UrlEncode(ByVal s As String) As String
    Dim b() As Byte, i As Long, c As Long, r As String, raw As String
    raw = Utf8Bytes(s)
    If LenB(raw) = 0 Then Exit Function
    b = raw
    For i = LBound(b) To UBound(b)
        c = b(i)
        If (c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or c = 45 Or c = 46 Or c = 95 Or c = 126 Then
            r = r & Chr$(c)
        Else
            r = r & "%" & Right$("0" & Hex$(c), 2)
        End If
    Next i
    UrlEncode = r
End Function


' =====================================================================
'  Background colour gallery: theme colours (with tints and shades), standard
'  colours, recent colours, "More Colors" and the eyedropper - like PowerPoint's own
' =====================================================================
Private Function ThemeBase(ByVal col As Long) As Long
    ' column order as in PowerPoint: Background 1, Text 1, Background 2, Text 2, Accent 1-6
    Dim idx As Variant, defs As Variant
    idx = Array(2, 1, 4, 3, 5, 6, 7, 8, 9, 10)
    defs = Array(&HFFFFFF, &H0&, &HE6E6E7, &H6A5444, &HC47244, &H317DED, &HA5A5A5, &HC0FF&, &HD59B5B, &H47AD70)
    ThemeBase = defs(col)
    On Error Resume Next
    ThemeBase = ActivePresentation.SlideMaster.Theme.ThemeColorScheme.Colors(idx(col)).RGB
End Function

' Office's tint / shade rows: lighter 80/60/40 %, darker 25/50 % (other steps for very light / dark colours).
Private Function ThemeTint(ByVal base As Long, ByVal row As Long, ByRef label As String) As Long
    Dim h As Double, s As Double, l As Double, p As Variant, f As Double, lighter As Boolean
    RgbToHsl base, h, s, l
    If l >= 0.999 Then
        p = Array(-0.05, -0.15, -0.25, -0.35, -0.5)
    ElseIf l <= 0.001 Then
        p = Array(0.5, 0.35, 0.25, 0.15, 0.05)
    ElseIf l > 0.8 Then
        p = Array(-0.1, -0.25, -0.5, -0.75, -0.9)
    ElseIf l < 0.2 Then
        p = Array(0.9, 0.75, 0.5, 0.25, 0.1)
    Else
        p = Array(0.8, 0.6, 0.4, -0.25, -0.5)
    End If
    f = p(row - 1)
    lighter = (f > 0)
    If lighter Then
        l = l * (1 - f) + f
        label = Tx("col.lighter", CLng(f * 100))
    Else
        l = l * (1 + f)
        label = Tx("col.darker", CLng(-f * 100))
    End If
    ThemeTint = HslToRgb(h, s, l)
End Function

Private Function GalColor(ByVal index As Long, Optional ByRef label As String) As Long
    Dim names As Variant, std As Variant, stdNames As Variant, rc As Variant, col As Long, row As Long, t As String
    names = Split(Tx("col.themeNames"), "|")
    std = Array(&HC0&, &HFF&, &HC0FF&, &HFFFF&, &H50D092, &H50B000, &HF0B000, &HC07000, &H602000, &HA03070)
    stdNames = Split(Tx("col.stdNames"), "|")
    If index < 60 Then
        col = index Mod 10
        row = index \ 10
        If row = 0 Then
            GalColor = ThemeBase(col)
            label = names(col)
        Else
            GalColor = ThemeTint(ThemeBase(col), row, t)
            label = Tx("col.tint", names(col), t)
        End If
    ElseIf index < 70 Then
        GalColor = std(index - 60)
        label = Tx("col.stdItem", stdNames(index - 60))
    Else
        rc = RecentColors()
        If index - 70 <= UBound(rc) Then
            GalColor = rc(index - 70)
            label = Tx("col.recentItem", ColorText(rc(index - 70)))
        End If
    End If
    If Len(label) > 0 And index < 70 Then label = Tx("col.withValue", label, ColorText(GalColor))
End Function

Private Function ColorText(ByVal c As Long) As String
    ColorText = "RGB " & (c And &HFF&) & "," & ((c \ &H100&) And &HFF&) & "," & ((c \ &H10000) And &HFF&)
End Function

Private Function RecentColors() As Variant
    Dim parts() As String, i As Long, n As Long, r() As Long, s As String
    s = GetSet("RecentColors", "")
    If Len(s) = 0 Then RecentColors = Array(): Exit Function
    parts = Split(s, ";")
    ReDim r(0 To UBound(parts))
    For i = 0 To UBound(parts)
        If Len(parts(i)) = 6 Then
            r(n) = RGB(CLng("&H" & Mid$(parts(i), 1, 2)), CLng("&H" & Mid$(parts(i), 3, 2)), CLng("&H" & Mid$(parts(i), 5, 2)))
            n = n + 1
        End If
    Next i
    If n = 0 Then RecentColors = Array(): Exit Function
    ReDim Preserve r(0 To n - 1)
    RecentColors = r
End Function

Private Sub AddRecentColor(ByVal c As Long)
    Dim rc As Variant, i As Long, s As String, n As Long, hx As String
    hx = Mid$(HexRGB(c), 2)
    s = hx
    n = 1
    rc = RecentColors()
    For i = 0 To UBound(rc)
        If rc(i) <> c And n < 10 Then
            s = s & ";" & Mid$(HexRGB(CLng(rc(i))), 2)
            n = n + 1
        End If
    Next i
    PutSet "RecentColors", s
End Sub

' Uses c as the custom background colour (and switches the background to "custom colour").
Private Sub SetBgColor(ByVal c As Long)
    PutSet "BgColor", (c And &HFF&) & "," & ((c \ &H100&) And &HFF&) & "," & ((c \ &H10000) And &HFF&)
    PutSet "BgType", CStr(BG_CUSTOM)
    AddRecentColor c
    RefreshRibbon
End Sub

' "More Colors": PowerPoint's own colour dialog (Windows: run on a temporary off-slide shape whose
' fill colour is read afterwards; macOS: the macOS colour panel, which is what PowerPoint uses there).
' When it cannot be opened, VecStamp's own colour panel takes over (then the system colour dialog,
' then a colour list as the last resort).
Private Sub DoMoreColors()
    Dim c As Long, trans As Double
    On Error GoTo EH
    c = BgColorValue(BG_CUSTOM)
#If Mac Then
    If MacColorPicker(c) Then SetBgColor c
#Else
    trans = BgAlpha() / 100#
    Select Case PickWithPowerPoint("ObjectFillMoreColorsDialog", False, c, trans)
        Case 1
            SetBgColor c
            PutSet "BgAlpha", CStr(CLng(trans * 100))
            RefreshRibbon
        Case -1
            Select Case ColorPanel(c)
                Case 1: SetBgColor c
                Case -1
                    If SysColorDialog(c, GetActiveWindow()) Then SetBgColor c
            End Select
    End Select
#End If
    Exit Sub
EH:
    If ColorListPanel(c) Then SetBgColor c
End Sub

Private Sub DoEyedropper()
    Dim c As Long, trans As Double
    c = BgColorValue(BG_CUSTOM)
    trans = 0
    Select Case PickWithPowerPoint("EyedropperFill", True, c, trans)
        Case 1
            SetBgColor c
        Case -1
            MsgBox Tx("msg.eyedropperView"), vbInformation, AppTitle()
    End Select
End Sub

' Returns 1 when a colour was picked, 0 when the user cancelled, -1 when the tool is not available.
Private Function PickWithPowerPoint(ByVal idMso As String, ByVal waitForPick As Boolean, _
                                    ByRef c As Long, ByRef trans As Double) As Long
    Dim win As Object, sld As Object, tmp As Object, prev As Object, pres As Object
    Dim c0 As Long, t0 As Double, started As Double, wasSaved As Long, changed As Boolean
    On Error Resume Next
    Set win = Application.ActiveWindow
    Set sld = win.View.Slide
    Set pres = win.Presentation
    PickWithPowerPoint = -1
    If sld Is Nothing Or pres Is Nothing Then Exit Function
    If win.Selection.Type = PP_SEL_SHAPES Then Set prev = win.Selection.ShapeRange
    Err.Clear
    On Error GoTo Failed
    wasSaved = pres.Saved
    Set tmp = sld.Shapes.AddShape(MSO_RECTANGLE, -120, 10, 60, 60)
    tmp.Name = "VecStamp colour"
    tmp.Line.Visible = MSO_FALSE
    tmp.Fill.Visible = MSO_TRUE
    tmp.Fill.Solid
    tmp.Fill.ForeColor.RGB = c
    tmp.Fill.Transparency = trans
    tmp.Select
    c0 = tmp.Fill.ForeColor.RGB
    t0 = tmp.Fill.Transparency
    If waitForPick Then
        ProgressBegin Tx("prog.eyedropper"), 1
        ProgressItem 1, ""
        ProgressAt 0, Tx("step.eyedropper")
    End If
    Application.CommandBars.ExecuteMso idMso
    If waitForPick Then
        started = Timer
        Do
            DoEvents
            If tmp.Fill.ForeColor.RGB <> c0 Then Exit Do
            If ProgressCancelled() Then Exit Do
            If Timer - started > 60 Or Timer < started Then Exit Do
            ProgressAt CDbl(Timer - started) / 60#, ""
        Loop
        ProgressEnd
    End If
    changed = (tmp.Fill.ForeColor.RGB <> c0) Or (Abs(tmp.Fill.Transparency - t0) > 0.001)
    If changed Then
        c = tmp.Fill.ForeColor.RGB
        trans = tmp.Fill.Transparency
    End If
    tmp.Delete
    pres.Saved = wasSaved
    On Error Resume Next
    If Not prev Is Nothing Then prev.Select
    PickWithPowerPoint = IIf(changed, 1, 0)
    Exit Function
Failed:
    Resume FailCleanup
FailCleanup:
    On Error Resume Next
    ProgressEnd
    If Not tmp Is Nothing Then
        tmp.Delete
        pres.Saved = wasSaved
    End If
    If Not prev Is Nothing Then prev.Select
    PickWithPowerPoint = -1
End Function

Private Sub RgbToHsl(ByVal c As Long, ByRef h As Double, ByRef s As Double, ByRef l As Double)
    Dim r As Double, g As Double, b As Double, mx As Double, mn As Double, d As Double
    r = (c And &HFF&) / 255#
    g = ((c \ &H100&) And &HFF&) / 255#
    b = ((c \ &H10000) And &HFF&) / 255#
    mx = r: If g > mx Then mx = g
    If b > mx Then mx = b
    mn = r: If g < mn Then mn = g
    If b < mn Then mn = b
    l = (mx + mn) / 2
    d = mx - mn
    If d < 0.000001 Then
        h = 0: s = 0
        Exit Sub
    End If
    If l > 0.5 Then s = d / (2 - mx - mn) Else s = d / (mx + mn)
    If mx = r Then
        h = (g - b) / d + IIf(g < b, 6, 0)
    ElseIf mx = g Then
        h = (b - r) / d + 2
    Else
        h = (r - g) / d + 4
    End If
    h = h / 6
End Sub

Private Function HslToRgb(ByVal h As Double, ByVal s As Double, ByVal l As Double) As Long
    Dim q As Double, p As Double
    If l < 0 Then l = 0
    If l > 1 Then l = 1
    If s < 0.000001 Then
        HslToRgb = RGB(CLng(l * 255), CLng(l * 255), CLng(l * 255))
        Exit Function
    End If
    If l < 0.5 Then q = l * (1 + s) Else q = l + s - l * s
    p = 2 * l - q
    HslToRgb = RGB(CLng(HueChannel(p, q, h + 1 / 3) * 255), CLng(HueChannel(p, q, h) * 255), CLng(HueChannel(p, q, h - 1 / 3) * 255))
End Function

Private Function HueChannel(ByVal p As Double, ByVal q As Double, ByVal t As Double) As Double
    If t < 0 Then t = t + 1
    If t > 1 Then t = t - 1
    If t < 1 / 6 Then
        HueChannel = p + (q - p) * 6 * t
    ElseIf t < 0.5 Then
        HueChannel = q
    ElseIf t < 2 / 3 Then
        HueChannel = p + (q - p) * (2 / 3 - t) * 6
    Else
        HueChannel = p
    End If
End Function

' Swatches for the ribbon (Windows: a small BMP file loaded as a picture).
Private Function SwatchPicture(ByVal c As Long, ByVal size As Long) As Object
    Set SwatchPicture = PatternPicture(0, c, size)
End Function

' Preview square of a background: transparent = checkerboard, grid = graph paper, else the colour.
Private Function BgPicture(ByVal bg As Long, ByVal size As Long) As Object
    Select Case bg
        Case BG_NONE: Set BgPicture = PatternPicture(1, RGB(255, 255, 255), size)
        Case BG_GRID: Set BgPicture = PatternPicture(2, RGB(255, 255, 255), size)
        Case Else: Set BgPicture = PatternPicture(0, BgColorValue(bg), size)
    End Select
End Function

' kind: 0 = solid colour c, 1 = checkerboard, 2 = grid on c; grey frame.
Private Function PatternPicture(ByVal kind As Long, ByVal c As Long, ByVal size As Long) As Object
#If Mac Then
#Else
    Dim p As String, b() As Byte, i As Long, x As Long, y As Long, rowBytes As Long, n As Long
    Dim r As Long, g As Long, bl As Long, cell As Long, raw As String
    On Error GoTo Done
    p = JoinPath(TempDirX(), "swatch" & kind & "_" & Mid$(HexRGB(c), 2) & "_" & size & ".bmp")
    If Not Fso().FileExists(p) Then
        rowBytes = ((size * 3 + 3) \ 4) * 4
        n = 54 + rowBytes * size
        ReDim b(0 To n - 1)
        BmpHeader b, size, size, rowBytes
        cell = size \ 4
        If cell < 2 Then cell = 2
        For y = 0 To size - 1
            For x = 0 To size - 1
                i = 54 + y * rowBytes + x * 3
                r = c And &HFF&: g = (c \ &H100&) And &HFF&: bl = (c \ &H10000) And &HFF&
                If kind = 1 And ((x \ cell) + (y \ cell)) Mod 2 = 0 Then r = 201: g = 206: bl = 214
                If kind = 2 And (x Mod cell = 0 Or y Mod cell = 0) Then r = 154: g = 165: bl = 180
                If x = 0 Or y = 0 Or x = size - 1 Or y = size - 1 Then r = 128: g = 128: bl = 128
                b(i) = bl: b(i + 1) = g: b(i + 2) = r
            Next x
        Next y
        raw = b
        WriteBytesX p, raw
    End If
    Set PatternPicture = LoadPicture(p)
Done:
#End If
End Function

Private Sub PutLong(ByRef b() As Byte, ByVal pos As Long, ByVal v As Long)
    b(pos) = v And &HFF&
    b(pos + 1) = (v \ &H100&) And &HFF&
    b(pos + 2) = (v \ &H10000) And &HFF&
    b(pos + 3) = (v \ &H1000000) And &HFF&
End Sub


' =====================================================================
'  Paths and names (pure string helpers, same on both platforms)
' =====================================================================
Private Function PathSep() As String
#If Mac Then
    PathSep = "/"
#Else
    PathSep = "\"
#End If
End Function

Private Function JoinPath(ByVal folder As String, ByVal fileName As String) As String
    If Right$(folder, 1) = PathSep() Then JoinPath = folder & fileName Else JoinPath = folder & PathSep() & fileName
End Function

Private Function ParentFolder(ByVal p As String) As String
    Dim k As Long
    k = InStrRev(p, PathSep())
    If k > 0 Then ParentFolder = Left$(p, k - 1)
End Function

Private Function FileNameOf(ByVal p As String) As String
    FileNameOf = Mid$(p, InStrRev(p, PathSep()) + 1)
End Function

Private Function ExtOf(ByVal p As String) As String
    Dim nm As String, k As Long
    nm = FileNameOf(p)
    k = InStrRev(nm, ".")
    If k > 0 Then ExtOf = Mid$(nm, k + 1)
End Function

Private Function BaseNameOf(ByVal p As String) As String
    Dim nm As String, k As Long
    nm = FileNameOf(p)
    k = InStrRev(nm, ".")
    If k > 1 Then BaseNameOf = Left$(nm, k - 1) Else BaseNameOf = nm
End Function

Private Function BaseNameFor(ByVal pres As Object, ByVal sr As Object) As String
    Dim s As String, idx As Long
    s = SafeName(BaseNameOf(pres.Name))
    On Error Resume Next
    idx = sr.Parent.SlideIndex
    On Error GoTo 0
    If idx > 0 Then s = s & "_" & Tx("file.page", idx)
    BaseNameFor = s
End Function

Private Function UniquePath(ByVal folder As String, ByVal baseStr As String, ByVal ext As String) As String
    Dim p As String, k As Long
    p = JoinPath(folder, baseStr & "." & ext)
    k = 2
    Do While FileExistsX(p)
        p = JoinPath(folder, baseStr & "_" & k & "." & ext)
        k = k + 1
    Loop
    UniquePath = p
End Function

Private Function SafeName(ByVal s As String) As String
    Dim bad As Variant
    For Each bad In Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab, vbCr, vbLf)
        s = Replace(s, CStr(bad), "_")
    Next
    s = Trim$(s)
    Do While Len(s) > 0 And Right$(s, 1) = "."
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(s) > 80 Then s = Left$(s, 80)
    If Len(s) = 0 Then s = "shape"
    SafeName = s
End Function

Private Function DefaultFolder(ByVal pres As Object) As String
    Dim f As String
    f = GetSet("Folder", "")
    If Len(f) > 0 Then
        If FolderExistsX(f) Then DefaultFolder = f: Exit Function
    End If
    f = ""
    On Error Resume Next
    f = pres.Path                                   ' may be a URL for cloud files
    On Error GoTo 0
    If Len(f) > 0 And InStr(f, "://") = 0 Then
        If FolderExistsX(f) Then DefaultFolder = f: Exit Function
    End If
    DefaultFolder = DesktopPathX()
End Function

Private Function InitialFolder(ByVal pres As Object) As String
    Dim f As String
    f = GetSet("LastFolder", "")
    If Len(f) > 0 Then
        If FolderExistsX(f) Then InitialFolder = f: Exit Function
    End If
    InitialFolder = DefaultFolder(pres)
End Function

Private Function FolderSettingText() As String
    Dim f As String
    f = GetSet("Folder", "")
    If Len(f) = 0 Then
        FolderSettingText = Tx("set.folderNone")
    Else
        FolderSettingText = f
    End If
End Function


' =====================================================================
'  Settings (Windows: HKCU\Software\VecStamp, macOS: VBA settings store)
' =====================================================================
Private Function GetSet(ByVal key As String, ByVal defVal As String) As String
    Dim v As Variant
    On Error Resume Next
#If Mac Then
    v = GetSetting(VS_NAME, "Settings", key, defVal)
#Else
    v = CreateObject("WScript.Shell").RegRead(REG_ROOT & key)
#End If
    If Err.Number <> 0 Then v = defVal
    On Error GoTo 0
    GetSet = CStr(v)
End Function

Private Sub PutSet(ByVal key As String, ByVal v As String)
    On Error Resume Next
#If Mac Then
    SaveSetting VS_NAME, "Settings", key, v
#Else
    CreateObject("WScript.Shell").RegWrite REG_ROOT & key, v, "REG_SZ"
#End If
End Sub

Private Sub ResetSettings()
    Dim k As Variant
    On Error Resume Next
    For Each k In Array("Format", "Mode", "Naming", "Dpi", "DpiCustom", "Background", "BgType", "BgColor", _
                        "BgAlpha", "GridStep", "Margin", "OpenFolder", "VisioEdit", "VsdxNative", "RecentColors", _
                        "Folder", "LastFolder")
#If Mac Then
        DeleteSetting VS_NAME, "Settings", CStr(k)
#Else
        CreateObject("WScript.Shell").RegDelete REG_ROOT & CStr(k)
#End If
    Next
End Sub

Private Function GetIdx(ByVal key As String, ByVal maxIdx As Long, Optional ByVal defIdx As Long = 0) As Long
    Dim v As String, r As Long
    v = GetSet(key, CStr(defIdx))
    r = defIdx
    If IsNumeric(v) Then r = CLng(v)
    If r < 0 Or r > maxIdx Then r = defIdx
    GetIdx = r
End Function

Private Function GetFlag(ByVal key As String, ByVal defVal As String) As Boolean
    GetFlag = (GetSet(key, defVal) = "1")
End Function

Private Function GetNum(ByVal key As String, ByVal defVal As Double, ByVal lo As Double, ByVal hi As Double) As Double
    Dim v As Double
    v = Val(GetSet(key, NumText(defVal)))
    If v < lo Or v > hi Then v = defVal
    GetNum = v
End Function

Private Function GetMargin() As Single
    GetMargin = GetNum("Margin", 0, 0, 500)
End Function

Private Function CustomDpi() As Long
    CustomDpi = CLng(GetNum("DpiCustom", 1200, 1, DPI_MAX))
End Function

Private Function DpiValue() As Long
    Select Case GetIdx("Dpi", 3, 1)
        Case 0: DpiValue = 150
        Case 2: DpiValue = 600
        Case 3: DpiValue = CustomDpi()
        Case Else: DpiValue = 300
    End Select
End Function

' Background type; migrates the v1.1 "white background" checkbox.
Private Function BgTypeIdx() As Long
    If Len(GetSet("BgType", "")) = 0 Then
        If GetFlag("Background", "0") Then BgTypeIdx = BG_WHITE Else BgTypeIdx = BG_NONE
    Else
        BgTypeIdx = GetIdx("BgType", BG_CUSTOM, BG_NONE)
    End If
End Function

Private Function BgAlpha() As Long
    BgAlpha = CLng(GetNum("BgAlpha", 0, 0, 100))
End Function

Private Function GridStep() As Single
    GridStep = GetNum("GridStep", 10, 2, 200)
End Function

Private Function BgColorValue(ByVal bg As Long) As Long
    Dim r As Long, g As Long, b As Long
    Select Case bg
        Case BG_BEIGE: BgColorValue = RGB(245, 240, 225)
        Case BG_CUSTOM
            If Not ParseColor(GetSet("BgColor", DEF_CUSTOM_COLOR), r, g, b) Then ParseColor DEF_CUSTOM_COLOR, r, g, b
            BgColorValue = RGB(r, g, b)
        Case Else: BgColorValue = RGB(255, 255, 255)
    End Select
End Function

Private Function ColorSettingText() As String
    Dim r As Long, g As Long, b As Long
    If Not ParseColor(GetSet("BgColor", DEF_CUSTOM_COLOR), r, g, b) Then ParseColor DEF_CUSTOM_COLOR, r, g, b
    ColorSettingText = r & "," & g & "," & b
End Function

' Accepts "255,248,235", "255 248 235", "rgb(255,248,235)", "#FFF8EB", "FFF8EB" or "#FFF".
Private Function ParseColor(ByVal s As String, ByRef r As Long, ByRef g As Long, ByRef b As Long) As Boolean
    Dim t As String, parts() As String, i As Long, n As Long, vals(0 To 2) As Long, ch As String
    t = LCase$(Trim$(s))
    t = Replace(t, "rgb", "")
    t = Replace(t, "(", "")
    t = Replace(t, ")", "")
    t = Replace(t, ChrW(&HFF0C), ",")               ' full-width comma
    If Left$(t, 1) = "#" Then t = Mid$(t, 2)
    If (Len(t) = 6 Or Len(t) = 3) And InStr(t, ",") = 0 And InStr(t, " ") = 0 Then
        For i = 1 To Len(t)
            ch = Mid$(t, i, 1)
            If InStr("0123456789abcdef", ch) = 0 Then Exit Function
        Next i
        If Len(t) = 3 Then t = Mid$(t, 1, 1) & Mid$(t, 1, 1) & Mid$(t, 2, 1) & Mid$(t, 2, 1) & Mid$(t, 3, 1) & Mid$(t, 3, 1)
        r = CLng("&H" & Mid$(t, 1, 2))
        g = CLng("&H" & Mid$(t, 3, 2))
        b = CLng("&H" & Mid$(t, 5, 2))
        ParseColor = True
        Exit Function
    End If
    t = Replace(t, ",", " ")
    t = Replace(t, ";", " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    parts = Split(Trim$(t), " ")
    If UBound(parts) <> 2 Then Exit Function
    For i = 0 To 2
        If Not IsNumeric(parts(i)) Then Exit Function
        n = CLng(Val(parts(i)))
        If n < 0 Or n > 255 Then Exit Function
        vals(i) = n
    Next i
    r = vals(0): g = vals(1): b = vals(2)
    ParseColor = True
End Function

' Parses a number typed into a ribbon edit box (accepts "," as decimal separator).
Private Function ReadNumber(ByVal s As String, ByVal lo As Double, ByVal hi As Double, ByRef v As Double) As Boolean
    s = Trim$(Replace(s, ",", "."))
    If Len(s) = 0 Then s = CStr(lo)
    If Not IsNumeric(Replace(s, ".", "")) Then Exit Function
    v = Val(s)
    If v < lo Or v > hi Then Exit Function
    ReadNumber = True
End Function

Private Function NumText(ByVal v As Double) As String
    NumText = Trim$(Str$(v))
    If Left$(NumText, 1) = "." Then NumText = "0" & NumText
End Function

' Number for XML (always "." as decimal separator).
Private Function NumXml(ByVal v As Double) As String
    Dim neg As Boolean, ip As Double, fp As Long, r As String
    v = Round(v, 6)
    If Abs(v) < 0.0000005 Then NumXml = "0": Exit Function
    neg = (v < 0)
    v = Abs(v)
    ip = Int(v)
    fp = CLng((v - ip) * 1000000#)
    If fp >= 1000000 Then ip = ip + 1: fp = fp - 1000000
    r = Trim$(Str$(ip))
    If fp > 0 Then
        r = r & "." & Right$("000000" & Trim$(Str$(fp)), 6)
        Do While Right$(r, 1) = "0"
            r = Left$(r, Len(r) - 1)
        Loop
    End If
    If neg Then r = "-" & r
    NumXml = r
End Function

Private Function MinS(ByVal a As Single, ByVal b As Single) As Single
    If a < b Then MinS = a Else MinS = b
End Function

Private Function MaxS(ByVal a As Single, ByVal b As Single) As Single
    If a > b Then MaxS = a Else MaxS = b
End Function


' =====================================================================
'  Formats and small UI helpers
' =====================================================================
Private Function FmtExt(ByVal fmt As Long) As String
    Select Case fmt
        Case FMT_EMZ: FmtExt = "emz"
        Case FMT_SVG: FmtExt = "svg"
        Case FMT_WMF: FmtExt = "wmf"
        Case FMT_PNG: FmtExt = "png"
        Case FMT_PDF: FmtExt = "pdf"
        Case FMT_VSDX: FmtExt = "vsdx"
        Case Else: FmtExt = "emf"
    End Select
End Function

Private Function FmtFromExt(ByVal ext As String) As Long
    Select Case ext
        Case "emf": FmtFromExt = FMT_EMF
        Case "emz": FmtFromExt = FMT_EMZ
        Case "svg": FmtFromExt = FMT_SVG
        Case "wmf": FmtFromExt = FMT_WMF
        Case "png": FmtFromExt = FMT_PNG
        Case "pdf": FmtFromExt = FMT_PDF
        Case "vsdx": FmtFromExt = FMT_VSDX
        Case Else: FmtFromExt = -1
    End Select
End Function

Private Function FilterLabel(ByVal fmt As Long) As String
    Select Case fmt
        Case 0 To FMT_LAST: FilterLabel = Tx("fmt." & FmtExt(fmt)) & " (*." & FmtExt(fmt) & ")"
        Case Else: FilterLabel = Tx("fmt.emf") & " (*.emf)"
    End Select
End Function

Private Function BgShortName(ByVal bg As Long) As String
    Select Case bg
        Case BG_WHITE: BgShortName = Tx("bg.white")
        Case BG_BEIGE: BgShortName = Tx("bg.beige")
        Case BG_GRID: BgShortName = Tx("bg.grid")
        Case BG_CUSTOM: BgShortName = Tx("bg.custom")
        Case Else: BgShortName = Tx("bg.none")
    End Select
End Function

Private Function BgName(ByVal bg As Long) As String
    Select Case bg
        Case BG_WHITE: BgName = Tx("bg.white")
        Case BG_BEIGE: BgName = Tx("bg.beige")
        Case BG_GRID: BgName = Tx("bg.gridValue", NumText(GridStep()))
        Case BG_CUSTOM: BgName = Tx("bg.customValue", ColorSettingText())
        Case Else: BgName = Tx("bg.none")
    End Select
End Function

Private Function SettingsSummary() As String
    Dim s As String, bg As Long
    s = Tx("sum.head") & vbCrLf & vbCrLf
    s = s & Tx("sum.format", UCase$(FmtExt(GetIdx("Format", FMT_LAST)))) & vbCrLf
    s = s & Tx("sum.mode", Tx(IIf(GetIdx("Mode", 1) = 1, "mode.split", "mode.merge"))) & vbCrLf
    s = s & Tx("sum.naming", Tx(IIf(GetIdx("Naming", 1) = 1, "naming.auto", "naming.ask"))) & vbCrLf
    s = s & Tx("sum.dpi", DpiValue()) & vbCrLf
    bg = BgTypeIdx()
    s = s & Tx("sum.bg", BgName(bg))
    If bg <> BG_NONE Then s = s & Tx("sum.alpha", BgAlpha())
    s = s & vbCrLf & Tx("sum.margin", NumText(GetMargin()))
    s = s & vbCrLf & Tx(IIf(GetFlag("VsdxNative", "1"), "sum.vsdxNative", "sum.vsdxPicture"))
    s = s & vbCrLf & Tx("sum.lang", LangName(CurLang()))
    SettingsSummary = s
End Function

Private Sub Report(ByVal files As Collection, ByVal showList As Boolean)
    Dim msg As String, i As Long
    If files.Count = 0 Then Exit Sub
    PutSet "LastFolder", ParentFolder(CStr(files(1)))
    RefreshRibbon
    If GetFlag("OpenFolder", "0") Then
        RevealFileX CStr(files(1))
        Exit Sub
    End If
    If Not showList Then Exit Sub
    msg = Tx("msg.exported", files.Count) & vbCrLf & vbCrLf
    For i = 1 To files.Count
        If i > 8 Then msg = msg & U("20262026") & vbCrLf: Exit For
        msg = msg & CStr(files(i)) & vbCrLf
    Next i
    MsgBox msg, vbInformation, AppTitle()
End Sub

' After Esc / Cancel: says so and lists the files finished before.
Private Sub ReportCancelled(ByVal files As Collection)
    Dim msg As String, i As Long
    If files.Count = 0 Then
        MsgBox Tx("msg.cancelled"), vbInformation, AppTitle()
        Exit Sub
    End If
    PutSet "LastFolder", ParentFolder(CStr(files(1)))
    msg = Tx("msg.cancelledPartial", files.Count) & vbCrLf & vbCrLf
    For i = 1 To files.Count
        If i > 8 Then msg = msg & U("20262026") & vbCrLf: Exit For
        msg = msg & CStr(files(i)) & vbCrLf
    Next i
    MsgBox msg, vbInformation, AppTitle()
End Sub

Private Sub RefreshRibbon()
    On Error Resume Next
    If Not gRibbon Is Nothing Then gRibbon.Invalidate
End Sub

Private Function AppTitle() As String
    AppTitle = Tx("app.title")
End Function


' =====================================================================
'  Interface languages: Simplified Chinese (default), Traditional Chinese, English, Japanese.
'  The texts live in src/i18n/strings.json; the build turns them into LoadStrings (generated
'  block below), which fills gStr with the texts of one language.
' =====================================================================
Private Function CurLang() As String
    If Len(gLang) = 0 Then
        gLang = GetSet("Lang", "zh-CN")
        If InStr("|zh-CN|zh-TW|en|ja|", "|" & gLang & "|") = 0 Then gLang = "zh-CN"
    End If
    CurLang = gLang
End Function

' Switches the interface language at once (the caller refreshes the ribbon).
Private Sub SetLang(ByVal code As String)
    If InStr("|zh-CN|zh-TW|en|ja|", "|" & code & "|") = 0 Then code = "zh-CN"
    PutSet "Lang", code
    gLang = code
    Set gStr = Nothing
End Sub

' Each language under its own name (not translated).
Private Function LangName(ByVal code As String) As String
    Select Case code
        Case "zh-TW": LangName = U("7E419AD44E2D6587")
        Case "en": LangName = "English"
        Case "ja": LangName = U("65E5672C8A9E")
        Case Else: LangName = U("7B804F534E2D6587")
    End Select
End Function

Private Sub EnsureStrings()
    If gStr Is Nothing Then
        Set gStr = New Collection
        LoadStrings CurLang()
    End If
End Sub

' The text of key in the current language with {0}, {1}, ... replaced by args (the key itself
' when it is missing). Err is preserved: Tx is often called inside error handlers.
Private Function Tx(ByVal key As String, ParamArray args() As Variant) As String
    Dim s As String, i As Long, en As Long, ed As String, es As String
    en = Err.Number: ed = Err.Description: es = Err.Source
    EnsureStrings
    On Error Resume Next
    s = gStr.Item(key)
    If Err.Number <> 0 Then s = key
    On Error GoTo 0
    For i = 0 To UBound(args)
        s = Replace(s, "{" & i & "}", CStr(args(i)))
    Next i
    If en <> 0 Then
        Err.Number = en
        Err.Description = ed
        Err.Source = es
    End If
    Tx = s
End Function

Private Function TxHas(ByVal key As String) As Boolean
    Dim s As String, en As Long, ed As String, es As String
    en = Err.Number: ed = Err.Description: es = Err.Source
    EnsureStrings
    On Error Resume Next
    s = gStr.Item(key)
    TxHas = (Err.Number = 0)
    On Error GoTo 0
    If en <> 0 Then
        Err.Number = en
        Err.Description = ed
        Err.Source = es
    End If
End Function

' Decodes UTF-16 code units written as 4-digit hex groups (used by the ASCII-only build).
Private Function U(ByVal h As String) As String
    Dim i As Long, s As String
    For i = 1 To Len(h) Step 4
        s = s & ChrW$(CLng("&H" & Mid$(h, i, 4)))
    Next i
    U = s
End Function


' >>> GENERATED: VSDX TEMPLATE PARTS, HELPER SCRIPT AND PRESET SHAPES (generated by build/build_bas.py - do not edit)
Private Function VsdxPart(ByVal partName As String) As String
    Dim s As String
    Select Case partName
        Case "content_types.xml"
            s = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types""><Default Extension=""emf"" ContentType=""image/x-emf"
            s = s & """/><Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/><Default Extension=""xml"" ContentType=""application/xml""/><Override PartName=""/vi"
            s = s & "sio/document.xml"" ContentType=""application/vnd.ms-visio.drawing.main+xml""/><Override PartName=""/visio/pages/pages.xml"" ContentType=""application/vnd.ms-visio.pages+xml""/><Override P"
            s = s & "artName=""/visio/pages/page1.xml"" ContentType=""application/vnd.ms-visio.page+xml""/><Override PartName=""/visio/windows.xml"" ContentType=""application/vnd.ms-visio.windows+xml""/><Overr"
            s = s & "ide PartName=""/docProps/core.xml"" ContentType=""application/vnd.openxmlformats-package.core-properties+xml""/><Override PartName=""/docProps/app.xml"" ContentType=""application/vnd.open"
            s = s & "xmlformats-officedocument.extended-properties+xml""/></Types>"
        Case "root.rels"
            s = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://sche"
            s = s & "mas.microsoft.com/visio/2010/relationships/document"" Target=""visio/document.xml""/><Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/package/2006/relationships/metadat"
            s = s & "a/core-properties"" Target=""docProps/core.xml""/><Relationship Id=""rId3"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"" Target=""docPro"
            s = s & "ps/app.xml""/></Relationships>"
        Case "app.xml"
            s = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"" xmlns:vt=""http://schemas.openxm"
            s = s & "lformats.org/officeDocument/2006/docPropsVTypes""><Application>VecStamp</Application><AppVersion>{VERSION}</AppVersion><HeadingPairs><vt:vector size=""2"" baseType=""variant""><vt:varia"
            s = s & "nt><vt:lpstr>Pages</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size=""1"" baseType=""lpstr""><vt:lpstr>Page-1</"
            s = s & "vt:lpstr></vt:vector></TitlesOfParts></Properties>"
        Case "core.xml"
            s = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"" xmlns:dc=""http://purl.o"
            s = s & "rg/dc/elements/1.1/"" xmlns:dcterms=""http://purl.org/dc/terms/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""><dc:title>{TITLE}</dc:title><dc:creator>VecStamp</dc:creator><d"
            s = s & "cterms:created xsi:type=""dcterms:W3CDTF"">{CREATED}</dcterms:created></cp:coreProperties>"
        Case "document.xml"
            s = "<?xml version='1.0' encoding='utf-8' ?><VisioDocument xmlns='http://schemas.microsoft.com/office/visio/2012/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/rel"
            s = s & "ationships' xml:space='preserve'><DocumentSettings TopPage='0' DefaultTextStyle='0' DefaultLineStyle='0' DefaultFillStyle='0' DefaultGuideStyle='0'><GlueSettings>9</GlueSettings><S"
            s = s & "napSettings>65847</SnapSettings><SnapExtensions>34</SnapExtensions><SnapAngles/><DynamicGridEnabled>1</DynamicGridEnabled><ProtectStyles>0</ProtectStyles><ProtectShapes>0</ProtectS"
            s = s & "hapes><ProtectMasters>0</ProtectMasters><ProtectBkgnds>0</ProtectBkgnds></DocumentSettings><Colors><ColorEntry IX='24' RGB='#7F7F7F'/><ColorEntry IX='25' RGB='#FFFFFF'/></Colors><F"
            s = s & "aceNames><FaceName NameU='Calibri' UnicodeRanges='-469750017 -1073732485 9 0' CharSets='536871423 0' Panose='2 15 5 2 2 2 4 3 2 4' Flags='325'/></FaceNames><StyleSheets><StyleSheet"
            s = s & " ID='0' NameU='No Style' IsCustomNameU='1' Name='No Style' IsCustomName='1'><Cell N='EnableLineProps' V='1'/><Cell N='EnableFillProps' V='1'/><Cell N='EnableTextProps' V='1'/><Cell"
            s = s & " N='HideForApply' V='0'/><Cell N='LineWeight' V='0.01041666666666667'/><Cell N='LineColor' V='0'/><Cell N='LinePattern' V='1'/><Cell N='Rounding' V='0'/><Cell N='EndArrowSize' V='2"
            s = s & "'/><Cell N='BeginArrow' V='0'/><Cell N='EndArrow' V='0'/><Cell N='LineCap' V='0'/><Cell N='BeginArrowSize' V='2'/><Cell N='LineColorTrans' V='0'/><Cell N='CompoundType' V='0'/><Cel"
            s = s & "l N='FillForegnd' V='1'/><Cell N='FillBkgnd' V='0'/><Cell N='FillPattern' V='1'/><Cell N='ShdwForegnd' V='0'/><Cell N='ShdwPattern' V='0'/><Cell N='FillForegndTrans' V='0'/><Cell N"
            s = s & "='FillBkgndTrans' V='0'/><Cell N='ShdwForegndTrans' V='0'/><Cell N='ShapeShdwType' V='0'/><Cell N='ShapeShdwOffsetX' V='0'/><Cell N='ShapeShdwOffsetY' V='0'/><Cell N='ShapeShdwObli"
            s = s & "queAngle' V='0'/><Cell N='ShapeShdwScaleFactor' V='1'/><Cell N='ShapeShdwBlur' V='0'/><Cell N='ShapeShdwShow' V='0'/><Cell N='LeftMargin' V='0'/><Cell N='RightMargin' V='0'/><Cell "
            s = s & "N='TopMargin' V='0'/><Cell N='BottomMargin' V='0'/><Cell N='VerticalAlign' V='1'/><Cell N='TextBkgnd' V='0'/><Cell N='DefaultTabStop' V='0.5905511811023622'/><Cell N='TextDirection"
            s = s & "' V='0'/><Cell N='TextBkgndTrans' V='0'/><Cell N='LockWidth' V='0'/><Cell N='LockHeight' V='0'/><Cell N='LockMoveX' V='0'/><Cell N='LockMoveY' V='0'/><Cell N='LockAspect' V='0'/><C"
            s = s & "ell N='LockDelete' V='0'/><Cell N='LockBegin' V='0'/><Cell N='LockEnd' V='0'/><Cell N='LockRotate' V='0'/><Cell N='LockCrop' V='0'/><Cell N='LockVtxEdit' V='0'/><Cell N='LockTextEd"
            s = s & "it' V='0'/><Cell N='LockFormat' V='0'/><Cell N='LockGroup' V='0'/><Cell N='LockCalcWH' V='0'/><Cell N='LockSelect' V='0'/><Cell N='LockCustProp' V='0'/><Cell N='LockFromGroupFormat"
            s = s & "' V='0'/><Cell N='LockThemeColors' V='0'/><Cell N='LockThemeEffects' V='0'/><Cell N='LockThemeConnectors' V='0'/><Cell N='LockThemeFonts' V='0'/><Cell N='LockThemeIndex' V='0'/><Ce"
            s = s & "ll N='LockReplace' V='0'/><Cell N='LockVariation' V='0'/><Cell N='NoObjHandles' V='0'/><Cell N='NonPrinting' V='0'/><Cell N='NoCtlHandles' V='0'/><Cell N='NoAlignBox' V='0'/><Cell "
            s = s & "N='UpdateAlignBox' V='0'/><Cell N='HideText' V='0'/><Cell N='DynFeedback' V='0'/><Cell N='GlueType' V='0'/><Cell N='WalkPreference' V='0'/><Cell N='BegTrigger' V='0' F='No Formula'"
            s = s & "/><Cell N='EndTrigger' V='0' F='No Formula'/><Cell N='ObjType' V='0'/><Cell N='Comment' V=''/><Cell N='IsDropSource' V='0'/><Cell N='NoLiveDynamics' V='0'/><Cell N='LocalizeMerge' "
            s = s & "V='0'/><Cell N='NoProofing' V='0'/><Cell N='Calendar' V='0'/><Cell N='LangID' V='en-US'/><Cell N='ShapeKeywords' V=''/><Cell N='DropOnPageScale' V='1'/><Cell N='TheData' V='0' F='N"
            s = s & "o Formula'/><Cell N='TheText' V='0' F='No Formula'/><Cell N='EventDblClick' V='0' F='No Formula'/><Cell N='EventXFMod' V='0' F='No Formula'/><Cell N='EventDrop' V='0' F='No Formula"
            s = s & "'/><Cell N='EventMultiDrop' V='0' F='No Formula'/><Cell N='HelpTopic' V=''/><Cell N='Copyright' V=''/><Cell N='LayerMember' V=''/><Cell N='XRulerDensity' V='32'/><Cell N='YRulerDen"
            s = s & "sity' V='32'/><Cell N='XRulerOrigin' V='0'/><Cell N='YRulerOrigin' V='0'/><Cell N='XGridDensity' V='8'/><Cell N='YGridDensity' V='8'/><Cell N='XGridSpacing' V='0'/><Cell N='YGridSp"
            s = s & "acing' V='0'/><Cell N='XGridOrigin' V='0'/><Cell N='YGridOrigin' V='0'/><Cell N='Gamma' V='1'/><Cell N='Contrast' V='0.5'/><Cell N='Brightness' V='0.5'/><Cell N='Sharpen' V='0'/><C"
            s = s & "ell N='Blur' V='0'/><Cell N='Denoise' V='0'/><Cell N='Transparency' V='0'/><Cell N='SelectMode' V='1'/><Cell N='DisplayMode' V='2'/><Cell N='IsDropTarget' V='0'/><Cell N='IsSnapTar"
            s = s & "get' V='1'/><Cell N='IsTextEditTarget' V='1'/><Cell N='DontMoveChildren' V='0'/><Cell N='ShapePermeableX' V='0'/><Cell N='ShapePermeableY' V='0'/><Cell N='ShapePermeablePlace' V='0"
            s = s & "'/><Cell N='Relationships' V='0'/><Cell N='ShapeFixedCode' V='0'/><Cell N='ShapePlowCode' V='0'/><Cell N='ShapeRouteStyle' V='0'/><Cell N='ShapePlaceStyle' V='0'/><Cell N='ConFixed"
            s = s & "Code' V='0'/><Cell N='ConLineJumpCode' V='0'/><Cell N='ConLineJumpStyle' V='0'/><Cell N='ConLineJumpDirX' V='0'/><Cell N='ConLineJumpDirY' V='0'/><Cell N='ShapePlaceFlip' V='0'/><C"
            s = s & "ell N='ConLineRouteExt' V='0'/><Cell N='ShapeSplit' V='0'/><Cell N='ShapeSplittable' V='0'/><Cell N='DisplayLevel' V='0'/><Cell N='ResizePage' V='0'/><Cell N='EnableGrid' V='0'/><C"
            s = s & "ell N='DynamicsOff' V='0'/><Cell N='CtrlAsInput' V='0'/><Cell N='AvoidPageBreaks' V='0'/><Cell N='PlaceStyle' V='0'/><Cell N='RouteStyle' V='0'/><Cell N='PlaceDepth' V='0'/><Cell N"
            s = s & "='PlowCode' V='0'/><Cell N='LineJumpCode' V='1'/><Cell N='LineJumpStyle' V='0'/><Cell N='PageLineJumpDirX' V='0'/><Cell N='PageLineJumpDirY' V='0'/><Cell N='LineToNodeX' V='0.09842"
            s = s & "519685039369'/><Cell N='LineToNodeY' V='0.09842519685039369'/><Cell N='BlockSizeX' V='0.1968503937007874'/><Cell N='BlockSizeY' V='0.1968503937007874'/><Cell N='AvenueSizeX' V='0.2"
            s = s & "952755905511811'/><Cell N='AvenueSizeY' V='0.2952755905511811'/><Cell N='LineToLineX' V='0.09842519685039369'/><Cell N='LineToLineY' V='0.09842519685039369'/><Cell N='LineJumpFacto"
            s = s & "rX' V='0.66666666666667'/><Cell N='LineJumpFactorY' V='0.66666666666667'/><Cell N='LineAdjustFrom' V='0'/><Cell N='LineAdjustTo' V='0'/><Cell N='PlaceFlip' V='0'/><Cell N='LineRout"
            s = s & "eExt' V='0'/><Cell N='PageShapeSplit' V='0'/><Cell N='PageLeftMargin' V='0.25'/><Cell N='PageRightMargin' V='0.25'/><Cell N='PageTopMargin' V='0.25'/><Cell N='PageBottomMargin' V='"
            s = s & "0.25'/><Cell N='ScaleX' V='1'/><Cell N='ScaleY' V='1'/><Cell N='PagesX' V='1'/><Cell N='PagesY' V='1'/><Cell N='CenterX' V='0'/><Cell N='CenterY' V='0'/><Cell N='OnPage' V='0'/><Ce"
            s = s & "ll N='PrintGrid' V='0'/><Cell N='PrintPageOrientation' V='1'/><Cell N='PaperKind' V='9'/><Cell N='PaperSource' V='7'/><Cell N='QuickStyleLineColor' V='100'/><Cell N='QuickStyleFill"
            s = s & "Color' V='100'/><Cell N='QuickStyleShadowColor' V='100'/><Cell N='QuickStyleFontColor' V='100'/><Cell N='QuickStyleLineMatrix' V='100'/><Cell N='QuickStyleFillMatrix' V='100'/><Cel"
            s = s & "l N='QuickStyleEffectsMatrix' V='100'/><Cell N='QuickStyleFontMatrix' V='100'/><Cell N='QuickStyleType' V='0'/><Cell N='QuickStyleVariation' V='0'/><Cell N='LineGradientDir' V='0'/"
            s = s & "><Cell N='LineGradientAngle' V='1.5707963267949'/><Cell N='FillGradientDir' V='0'/><Cell N='FillGradientAngle' V='1.5707963267949'/><Cell N='LineGradientEnabled' V='0'/><Cell N='Fi"
            s = s & "llGradientEnabled' V='0'/><Cell N='RotateGradientWithShape' V='1'/><Cell N='UseGroupGradient' V='0'/><Cell N='BevelTopType' V='0'/><Cell N='BevelTopWidth' V='0'/><Cell N='BevelTopH"
            s = s & "eight' V='0'/><Cell N='BevelBottomType' V='0'/><Cell N='BevelBottomWidth' V='0'/><Cell N='BevelBottomHeight' V='0'/><Cell N='BevelDepthColor' V='1'/><Cell N='BevelDepthSize' V='0'/"
            s = s & "><Cell N='BevelContourColor' V='0'/><Cell N='BevelContourSize' V='0'/><Cell N='BevelMaterialType' V='0'/><Cell N='BevelLightingType' V='0'/><Cell N='BevelLightingAngle' V='0'/><Cel"
            s = s & "l N='RotationXAngle' V='0'/><Cell N='RotationYAngle' V='0'/><Cell N='RotationZAngle' V='0'/><Cell N='RotationType' V='0'/><Cell N='Perspective' V='0'/><Cell N='DistanceFromGround' "
            s = s & "V='0'/><Cell N='KeepTextFlat' V='0'/><Cell N='ReflectionTrans' V='0'/><Cell N='ReflectionSize' V='0'/><Cell N='ReflectionDist' V='0'/><Cell N='ReflectionBlur' V='0'/><Cell N='GlowC"
            s = s & "olor' V='1'/><Cell N='GlowColorTrans' V='0'/><Cell N='GlowSize' V='0'/><Cell N='SoftEdgesSize' V='0'/><Cell N='SketchSeed' V='0'/><Cell N='SketchEnabled' V='0'/><Cell N='SketchAmou"
            s = s & "nt' V='5'/><Cell N='SketchLineWeight' V='0.04166666666666666' U='PT'/><Cell N='SketchLineChange' V='0.14'/><Cell N='SketchFillChange' V='0.1'/><Cell N='ColorSchemeIndex' V='0'/><Ce"
            s = s & "ll N='EffectSchemeIndex' V='0'/><Cell N='ConnectorSchemeIndex' V='0'/><Cell N='FontSchemeIndex' V='0'/><Cell N='ThemeIndex' V='0'/><Cell N='VariationColorIndex' V='0'/><Cell N='Var"
            s = s & "iationStyleIndex' V='0'/><Cell N='EmbellishmentIndex' V='0'/><Cell N='ReplaceLockShapeData' V='0'/><Cell N='ReplaceLockText' V='0'/><Cell N='ReplaceLockFormat' V='0'/><Cell N='Repl"
            s = s & "aceCopyCells' V='0' U='BOOL' F='No Formula'/><Cell N='PageWidth' V='0' F='No Formula'/><Cell N='PageHeight' V='0' F='No Formula'/><Cell N='ShdwOffsetX' V='0' F='No Formula'/><Cell "
            s = s & "N='ShdwOffsetY' V='0' F='No Formula'/><Cell N='PageScale' V='0' U='MM' F='No Formula'/><Cell N='DrawingScale' V='0' U='MM' F='No Formula'/><Cell N='DrawingSizeType' V='0' F='No For"
            s = s & "mula'/><Cell N='DrawingScaleType' V='0' F='No Formula'/><Cell N='InhibitSnap' V='0' F='No Formula'/><Cell N='PageLockReplace' V='0' U='BOOL' F='No Formula'/><Cell N='PageLockDuplic"
            s = s & "ate' V='0' U='BOOL' F='No Formula'/><Cell N='UIVisibility' V='0' F='No Formula'/><Cell N='ShdwType' V='0' F='No Formula'/><Cell N='ShdwObliqueAngle' V='0' F='No Formula'/><Cell N='"
            s = s & "ShdwScaleFactor' V='0' F='No Formula'/><Cell N='DrawingResizeType' V='0' F='No Formula'/><Section N='Character'><Row IX='0'><Cell N='Font' V='Calibri'/><Cell N='Color' V='0'/><Cell"
            s = s & " N='Style' V='0'/><Cell N='Case' V='0'/><Cell N='Pos' V='0'/><Cell N='FontScale' V='1'/><Cell N='Size' V='0.1666666666666667'/><Cell N='DblUnderline' V='0'/><Cell N='Overline' V='0"
            s = s & "'/><Cell N='Strikethru' V='0'/><Cell N='DoubleStrikethrough' V='0'/><Cell N='Letterspace' V='0'/><Cell N='ColorTrans' V='0'/><Cell N='AsianFont' V='0'/><Cell N='ComplexScriptFont' "
            s = s & "V='0'/><Cell N='ComplexScriptSize' V='-1'/><Cell N='LangID' V='en-US'/></Row></Section><Section N='Paragraph'><Row IX='0'><Cell N='IndFirst' V='0'/><Cell N='IndLeft' V='0'/><Cell N"
            s = s & "='IndRight' V='0'/><Cell N='SpLine' V='-1.2'/><Cell N='SpBefore' V='0'/><Cell N='SpAfter' V='0'/><Cell N='HorzAlign' V='1'/><Cell N='Bullet' V='0'/><Cell N='BulletStr' V=''/><Cell "
            s = s & "N='BulletFont' V='0'/><Cell N='BulletFontSize' V='-1'/><Cell N='TextPosAfterBullet' V='0'/><Cell N='Flags' V='0'/></Row></Section><Section N='Tabs'><Row IX='0'/></Section><Section "
            s = s & "N='LineGradient'><Row IX='0'><Cell N='GradientStopColor' V='1'/><Cell N='GradientStopColorTrans' V='0'/><Cell N='GradientStopPosition' V='0'/></Row></Section><Section N='FillGradie"
            s = s & "nt'><Row IX='0'><Cell N='GradientStopColor' V='1'/><Cell N='GradientStopColorTrans' V='0'/><Cell N='GradientStopPosition' V='0'/></Row></Section></StyleSheet></StyleSheets><Documen"
            s = s & "tSheet NameU='TheDoc' IsCustomNameU='1' Name='TheDoc' IsCustomName='1' LineStyle='0' FillStyle='0' TextStyle='0'><Cell N='OutputFormat' V='0'/><Cell N='LockPreview' V='0'/><Cell N="
            s = s & "'AddMarkup' V='0'/><Cell N='ViewMarkup' V='0'/><Cell N='DocLockReplace' V='0' U='BOOL'/><Cell N='NoCoauth' V='0' U='BOOL'/><Cell N='DocLockDuplicatePage' V='0' U='BOOL'/><Cell N='P"
            s = s & "reviewQuality' V='0'/><Cell N='PreviewScope' V='0'/><Cell N='DocLangID' V='en-US'/><Section N='User'><Row N='msvNoAutoConnect'><Cell N='Value' V='1'/><Cell N='Prompt' V='' F='No Fo"
            s = s & "rmula'/></Row></Section></DocumentSheet></VisioDocument>"
        Case "document.xml.rels"
            s = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://sche"
            s = s & "mas.microsoft.com/visio/2010/relationships/pages"" Target=""pages/pages.xml""/><Relationship Id=""rId2"" Type=""http://schemas.microsoft.com/visio/2010/relationships/windows"" Target=""win"
            s = s & "dows.xml""/></Relationships>"
        Case "windows.xml"
            s = "<?xml version='1.0' encoding='utf-8' ?><Windows ClientWidth='1280' ClientHeight='720' xmlns='http://schemas.microsoft.com/office/visio/2012/main' xmlns:r='http://schemas.openxmlfor"
            s = s & "mats.org/officeDocument/2006/relationships' xml:space='preserve'><Window ID='0' WindowType='Drawing' WindowState='1073741824' ContainerType='Page' Page='0' ViewScale='-1' ViewCente"
            s = s & "rX='{CX}' ViewCenterY='{CY}'><ShowRulers>1</ShowRulers><ShowGrid>1</ShowGrid><ShowPageBreaks>0</ShowPageBreaks><ShowGuides>1</ShowGuides><ShowConnectionPoints>1</ShowConnectionPoin"
            s = s & "ts><GlueSettings>9</GlueSettings><SnapSettings>65847</SnapSettings><SnapExtensions>34</SnapExtensions><SnapAngles/><DynamicGridEnabled>1</DynamicGridEnabled><TabSplitterPos>0.5</Ta"
            s = s & "bSplitterPos></Window></Windows>"
        Case "pages.xml"
            s = "<?xml version='1.0' encoding='utf-8' ?><Pages xmlns='http://schemas.microsoft.com/office/visio/2012/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationshi"
            s = s & "ps' xml:space='preserve'><Page ID='0' NameU='Page-1' Name='Page-1' ViewScale='-1' ViewCenterX='{CX}' ViewCenterY='{CY}'><PageSheet LineStyle='0' FillStyle='0' TextStyle='0'><Cell N"
            s = s & "='PageWidth' V='{PW}'/><Cell N='PageHeight' V='{PH}'/><Cell N='ShdwOffsetX' V='0.125'/><Cell N='ShdwOffsetY' V='-0.125'/><Cell N='PageScale' V='1' U='IN'/><Cell N='DrawingScale' V="
            s = s & "'1' U='IN'/><Cell N='DrawingSizeType' V='3'/><Cell N='DrawingScaleType' V='0'/><Cell N='InhibitSnap' V='0'/><Cell N='UIVisibility' V='0'/><Cell N='ShdwType' V='0'/><Cell N='ShdwObl"
            s = s & "iqueAngle' V='0'/><Cell N='ShdwScaleFactor' V='1'/><Cell N='DrawingResizeType' V='1'/><Cell N='PageShapeSplit' V='1'/></PageSheet><Rel r:id='rId1'/></Page></Pages>"
        Case "pages.xml.rels"
            s = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://sche"
            s = s & "mas.microsoft.com/visio/2010/relationships/page"" Target=""page1.xml""/></Relationships>"
        Case "page1.xml"
            s = "<?xml version='1.0' encoding='utf-8' ?><PageContents xmlns='http://schemas.microsoft.com/office/visio/2012/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/rela"
            s = s & "tionships' xml:space='preserve'><Shapes>{SHAPES}</Shapes></PageContents>"
        Case "page1.xml.rels"
            s = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">{RELS}</Relationships>"
    End Select
    VsdxPart = s
End Function

Private Function VsdxMap() As Variant
    VsdxMap = Array("content_types.xml=[Content_Types].xml", _
                    "root.rels=_rels/.rels", _
                    "app.xml=docProps/app.xml", _
                    "core.xml=docProps/core.xml", _
                    "document.xml=visio/document.xml", _
                    "document.xml.rels=visio/_rels/document.xml.rels", _
                    "windows.xml=visio/windows.xml", _
                    "pages.xml=visio/pages/pages.xml", _
                    "pages.xml.rels=visio/pages/_rels/pages.xml.rels", _
                    "page1.xml=visio/pages/page1.xml", _
                    "page1.xml.rels=visio/pages/_rels/page1.xml.rels")
End Function

Private Function HelperScript1() As String
    Dim s As String
    s = s & "# VecStamp (Shi Yin) - XPS page -> SVG converter." & vbLf
    s = s & "#" & vbLf
    s = s & "# Used by the PowerPoint add-in on PowerPoint versions that cannot export SVG themselves" & vbLf
    s = s & "# (Windows, PowerPoint 2013 - 2021 / LTSC): the add-in prints the drawing to XPS with" & vbLf
    s = s & "# PowerPoint (""ExportAsFixedFormat"") and this script turns the first page into SVG," & vbLf
    s = s & "# cropped to the drawing. Paths, gradients, images, clips and transforms are converted" & vbLf
    s = s & "# 1:1; text stays text (font family, size, exact glyph positions) and the subset fonts" & vbLf
    s = s & "# embedded in the XPS are embedded in the SVG as well." & vbLf
    s = s & "#" & vbLf
    s = s & "# Windows PowerShell 5.1 and PowerShell 7 compatible; no modules needed." & vbLf
    s = s & "# The add-in embeds this file (build/build_bas.py), writes it to its temp folder and runs:" & vbLf
    s = s & "#   & ([scriptblock]::Create((Get-Content -Raw xps2svg.ps1))) -Xps in.xps -Out out.svg -Frame ""l,t,w,h"" -Slide ""w,h""" & vbLf
    s = s & "# Frame and Slide are in points (slide coordinates). Exit code 0 = success, 1 = failure," & vbLf
    s = s & "# 3 = gave up after -MaxSeconds. With -Progress the script keeps ""<percent>|<step>|<a>|<b>"" in that" & vbLf
    s = s & "# file (steps: open, page, elem, font, write, done) so the add-in can show what is happening." & vbLf
    s = s & "#" & vbLf
    s = s & "# Copyright (c) 2026 vluckyzhang <vluckyzhang@gmail.com>. MIT License." & vbLf
    s = s & "" & vbLf
    s = s & "param(" & vbLf
    s = s & "    [Parameter(Mandatory = $true)][string]$Xps," & vbLf
    s = s & "    [Parameter(Mandatory = $true)][string]$Out," & vbLf
    s = s & "    [string]$Frame = ''," & vbLf
    s = s & "    [string]$Slide = ''," & vbLf
    s = s & "    [string]$Progress = ''," & vbLf
    s = s & "    [int]$MaxSeconds = 150," & vbLf
    s = s & "    [int]$MaxFontBytes = 6000000," & vbLf
    s = s & "    [switch]$NoEmbedFonts" & vbLf
    s = s & ")" & vbLf
    s = s & "" & vbLf
    s = s & "$ErrorActionPreference = 'Stop'" & vbLf
    s = s & "Add-Type -AssemblyName System.IO.Compression" & vbLf
    s = s & "Add-Type -AssemblyName System.IO.Compression.FileSystem" & vbLf
    s = s & "$Inv = [Globalization.CultureInfo]::InvariantCulture" & vbLf
    s = s & "$KeyNs = 'http://schemas.microsoft.com/xps/2005/06/resourcedictionary-key'" & vbLf
    s = s & "$Watch = [Diagnostics.Stopwatch]::StartNew()" & vbLf
    s = s & "$script:LastReport = -1.0" & vbLf
    s = s & "$script:LastPct = -1" & vbLf
    s = s & "$script:Done = 0" & vbLf
    s = s & "$script:Total = 1" & vbLf
    s = s & "" & vbLf
    s = s & "# ------------------------------------------------------------------ progress and time limit" & vbLf
    s = s & "function Report([int]$pct, [string]$step, [int]$a = 0, [int]$b = 0) {" & vbLf
    s = s & "    if (-not $Progress) { return }" & vbLf
    s = s & "    $now = $Watch.Elapsed.TotalMilliseconds" & vbLf
    s = s & "    if ($pct -eq $script:LastPct -and $now - $script:LastReport -lt 400) { return }" & vbLf
    s = s & "    if ($step -eq 'elem' -and $now - $script:LastReport -lt 150) { return }" & vbLf
    s = s & "    $script:LastPct = $pct" & vbLf
    s = s & "    $script:LastReport = $now" & vbLf
    s = s & "    try { [IO.File]::WriteAllText($Progress, ('{0}|{1}|{2}|{3}' -f $pct, $step, $a, $b)) } catch { }" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function CheckTime {" & vbLf
    s = s & "    if ($Watch.Elapsed.TotalSeconds -gt $MaxSeconds) { throw 'timeout' }" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function Tick {" & vbLf
    s = s & "    $script:Done++" & vbLf
    s = s & "    if (($script:Done % 8) -eq 0) {" & vbLf
    s = s & "        CheckTime" & vbLf
    s = s & "        Report ([int](15 + 70 * [Math]::Min(1.0, $script:Done / [double]$script:Total))) 'elem' $script:Done $script:Total" & vbLf
    s = s & "    }" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "# ------------------------------------------------------------------ helpers" & vbLf
    s = s & "function N([double]$v) {" & vbLf
    s = s & "    if ([Math]::Abs($v) -lt 0.0005) { return '0' }" & vbLf
    s = s & "    return $v.ToString('0.###', $Inv)" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function Num([string]$s) {" & vbLf
    s = s & "    return [double]::Parse($s.Trim(), [Globalization.NumberStyles]::Float, $Inv)" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function Esc([string]$s) {" & vbLf
    s = s & "    $sb = New-Object System.Text.StringBuilder" & vbLf
    s = s & "    foreach ($ch in $s.ToCharArray()) {" & vbLf
    s = s & "        $c = [int]$ch" & vbLf
    s = s & "        if ($c -eq 38) { [void]$sb.Append('&amp;') }" & vbLf
    s = s & "        elseif ($c -eq 60) { [void]$sb.Append('&lt;') }" & vbLf
    s = s & "        elseif ($c -eq 62) { [void]$sb.Append('&gt;') }" & vbLf
    s = s & "        elseif ($c -eq 34) { [void]$sb.Append('&quot;') }" & vbLf
    s = s & "        elseif ($c -lt 32 -and $c -ne 9 -and $c -ne 10 -and $c -ne 13) { }" & vbLf
    s = s & "        elseif ($c -eq 0xFFFE -or $c -eq 0xFFFF) { }" & vbLf
    s = s & "        else { [void]$sb.Append($ch) }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return $sb.ToString()" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function Numbers([string]$s) {" & vbLf
    s = s & "    $list = New-Object System.Collections.Generic.List[double]" & vbLf
    s = s & "    foreach ($m in [regex]::Matches($s, '[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?')) { $list.Add((Num $m.Value)) }" & vbLf
    s = s & "    return , $list.ToArray()" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "# 2D affine matrices as [a, b, c, d, e, f] (x' = a x + c y + e, y' = b x + d y + f), XPS / SVG order" & vbLf
    s = s & "function MatMul($m, $n) {" & vbLf
    s = s & "    # result = apply m first, then n" & vbLf
    s = s & "    return @(($m[0] * $n[0] + $m[1] * $n[2]), ($m[0] * $n[1] + $m[1] * $n[3])," & vbLf
    s = s & "        ($m[2] * $n[0] + $m[3] * $n[2]), ($m[2] * $n[1] + $m[3] * $n[3])," & vbLf
    s = s & "        ($m[4] * $n[0] + $m[5] * $n[2] + $n[4]), ($m[4] * $n[1] + $m[5] * $n[3] + $n[5]))" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function MatText($m) {" & vbLf
    s = s & "    return 'matrix(' + (($m | ForEach-Object { N $_ }) -join ' ') + ')'" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function ParseMatrix([string]$s) {" & vbLf
    s = s & "    $v = Numbers $s" & vbLf
    s = s & "    if ($v.Count -ne 6) { return $null }" & vbLf
    s = s & "    return @($v[0], $v[1], $v[2], $v[3], $v[4], $v[5])" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "# ------------------------------------------------------------------ package" & vbLf
    s = s & "$Zip = [IO.Compression.ZipFile]::OpenRead($Xps)" & vbLf
    s = s & "$Parts = @{}" & vbLf
    s = s & "foreach ($e in $Zip.Entries) {" & vbLf
    s = s & "    $name = '/' + [Uri]::UnescapeDataString($e.FullName.Replace('\', '/'))" & vbLf
    s = s & "    $name = [regex]::Replace($name, '/\[\d+\]\.(last\.)?piece$', '')" & vbLf
    s = s & "    $Parts[$name.ToLowerInvariant()] = $e" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function ReadBytes([string]$part) {" & vbLf
    s = s & "    $e = $Parts[$part.ToLowerInvariant()]" & vbLf
    s = s & "    if ($null -eq $e) { return $null }" & vbLf
    s = s & "    $s = $e.Open()" & vbLf
    s = s & "    $ms = New-Object IO.MemoryStream" & vbLf
    s = s & "    $s.CopyTo($ms)" & vbLf
    s = s & "    $s.Dispose()" & vbLf
    s = s & "    return , $ms.ToArray()" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function ReadXml([string]$part) {" & vbLf
    s = s & "    $b = ReadBytes $part" & vbLf
    s = s & "    if ($null -eq $b) { return $null }" & vbLf
    s = s & "    $x = New-Object System.Xml.XmlDocument" & vbLf
    s = s & "    $x.PreserveWhitespace = $true" & vbLf
    s = s & "    $x.Load((New-Object IO.MemoryStream (, $b)))" & vbLf
    s = s & "    return $x" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function Resolve([string]$base, [string]$uri) {" & vbLf
    s = s & "    $u = [Uri]::UnescapeDataString($uri.Trim())" & vbLf
    s = s & "    if ($u.StartsWith('/')) { return $u }" & vbLf
    s = s & "    $dir = $base.Substring(0, $base.LastIndexOf('/') + 1)" & vbLf
    s = s & "    $stack = New-Object System.Collections.ArrayList" & vbLf
    s = s & "    foreach ($seg in ($dir + $u).Split('/')) {" & vbLf
    s = s & "        if ($seg -eq '' -or $seg -eq '.') { continue }" & vbLf
    s = s & "        if ($seg -eq '..') { if ($stack.Count) { $stack.RemoveAt($stack.Count - 1) } ; continue }" & vbLf
    s = s & "        [void]$stack.Add($seg)" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return '/' + ($stack -join '/')" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function FirstPage {" & vbLf
    s = s & "    $rels = ReadXml '/_rels/.rels'" & vbLf
    s = s & "    if ($rels) {" & vbLf
    s = s & "        foreach ($r in $rels.DocumentElement.ChildNodes) {" & vbLf
    s = s & "            if ($r.LocalName -eq 'Relationship' -and $r.GetAttribute('Type') -match 'fixedrepresentation$') {" & vbLf
    s = s & "                $seqPart = Resolve '/' $r.GetAttribute('Target')" & vbLf
    s = s & "                $seq = ReadXml $seqPart" & vbLf
    s = s & "                if (-not $seq) { break }" & vbLf
    s = s & "                foreach ($d in $seq.DocumentElement.ChildNodes) {" & vbLf
    s = s & "                    if ($d.LocalName -ne 'DocumentReference') { continue }" & vbLf
    s = s & "                    $docPart = Resolve $seqPart $d.GetAttribute('Source')" & vbLf
    s = s & "                    $doc = ReadXml $docPart" & vbLf
    s = s & "                    if (-not $doc) { continue }" & vbLf
    s = s & "                    foreach ($p in $doc.DocumentElement.ChildNodes) {" & vbLf
    s = s & "                        if ($p.LocalName -eq 'PageContent') { return (Resolve $docPart $p.GetAttribute('Source')) }" & vbLf
    s = s & "                    }" & vbLf
    s = s & "                }" & vbLf
    s = s & "            }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    $names = @($Parts.Keys | Where-Object { $_ -like '*.fpage' } | Sort-Object)" & vbLf
    s = s & "    if ($names.Count) { return $names[0] }" & vbLf
    s = s & "    throw 'no FixedPage in XPS'" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "# ------------------------------------------------------------------ output state" & vbLf
    s = s & "$Defs = New-Object System.Text.StringBuilder" & vbLf
    s = s & "$Body = New-Object System.Text.StringBuilder" & vbLf
    s = s & "$script:IdN = 0" & vbLf
    s = s & "$Fonts = @{}          # font part -> info hashtable" & vbLf
    s = s & "$Resources = @{}      # x:Key -> XmlElement" & vbLf
    s = s & "$script:Box = $null   # content bounding box in page units [x0, y0, x1, y1]" & vbLf
    s = s & "$script:PageW = 0.0" & vbLf
    s = s & "$script:PageH = 0.0" & vbLf
    s = s & "$script:FontsEmbedded = 0" & vbLf
    s = s & "" & vbLf
    s = s & "function NewId([string]$prefix) {" & vbLf
    s = s & "    $script:IdN++" & vbLf
    s = s & "    return $prefix + $script:IdN" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function AddBox($m, [double[]]$pts, [double]$pad) {" & vbLf
    s = s & "    # pts: x0, y0, x1, y1, ... in local coordinates; m: matrix to page coordinates" & vbLf
    s = s & "    for ($i = 0; $i + 1 -lt $pts.Count; $i += 2) {" & vbLf
    s = s & "        $x = $m[0] * $pts[$i] + $m[2] * $pts[$i + 1] + $m[4]" & vbLf
    s = s & "        $y = $m[1] * $pts[$i] + $m[3] * $pts[$i + 1] + $m[5]" & vbLf
    s = s & "        if ($null -eq $script:Box) { $script:Box = @(($x - $pad), ($y - $pad), ($x + $pad), ($y + $pad)) }" & vbLf
    HelperScript1 = s
End Function

Private Function HelperScript2() As String
    Dim s As String
    s = s & "        else {" & vbLf
    s = s & "            if ($x - $pad -lt $script:Box[0]) { $script:Box[0] = $x - $pad }" & vbLf
    s = s & "            if ($y - $pad -lt $script:Box[1]) { $script:Box[1] = $y - $pad }" & vbLf
    s = s & "            if ($x + $pad -gt $script:Box[2]) { $script:Box[2] = $x + $pad }" & vbLf
    s = s & "            if ($y + $pad -gt $script:Box[3]) { $script:Box[3] = $y + $pad }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "# ------------------------------------------------------------------ geometry" & vbLf
    s = s & "function FigureBounds([string]$d) {" & vbLf
    s = s & "    # absolute points of an abbreviated-syntax path (end and control points) for bounding boxes" & vbLf
    s = s & "    $pts = New-Object System.Collections.Generic.List[double]" & vbLf
    s = s & "    $toks = [regex]::Matches($d, '[A-Za-z]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?')" & vbLf
    s = s & "    $cmd = 'M'; $x = 0.0; $y = 0.0; $sx = 0.0; $sy = 0.0" & vbLf
    s = s & "    $nums = New-Object System.Collections.Generic.List[double]" & vbLf
    s = s & "    $flush = {" & vbLf
    s = s & "        param($c, $vals)" & vbLf
    s = s & "        $rel = ($c -cmatch '[a-z]')" & vbLf
    s = s & "        $C = $c.ToUpperInvariant()" & vbLf
    s = s & "        $k = 0" & vbLf
    s = s & "        $arity = @{ 'M' = 2; 'L' = 2; 'H' = 1; 'V' = 1; 'C' = 6; 'Q' = 4; 'S' = 4; 'T' = 2; 'A' = 7; 'Z' = 0 }[$C]" & vbLf
    s = s & "        if ($null -eq $arity -or $arity -eq 0) { return }" & vbLf
    s = s & "        while ($k + $arity -le $vals.Count) {" & vbLf
    s = s & "            switch ($C) {" & vbLf
    s = s & "                'H' { $script:px = $(if ($rel) { $script:px + $vals[$k] } else { $vals[$k] }) }" & vbLf
    s = s & "                'V' { $script:py = $(if ($rel) { $script:py + $vals[$k] } else { $vals[$k] }) }" & vbLf
    s = s & "                'A' {" & vbLf
    s = s & "                    $nx = $vals[$k + 5]; $ny = $vals[$k + 6]" & vbLf
    s = s & "                    if ($rel) { $nx += $script:px; $ny += $script:py }" & vbLf
    s = s & "                    $r = [Math]::Max([Math]::Abs($vals[$k]), [Math]::Abs($vals[$k + 1]))" & vbLf
    s = s & "                    $pts.Add($script:px - $r); $pts.Add($script:py - $r); $pts.Add($script:px + $r); $pts.Add($script:py + $r)" & vbLf
    s = s & "                    $script:px = $nx; $script:py = $ny" & vbLf
    s = s & "                }" & vbLf
    s = s & "                default {" & vbLf
    s = s & "                    for ($j = 0; $j -lt $arity; $j += 2) {" & vbLf
    s = s & "                        $nx = $vals[$k + $j]; $ny = $vals[$k + $j + 1]" & vbLf
    s = s & "                        if ($rel) { $nx += $script:px; $ny += $script:py }" & vbLf
    s = s & "                        $pts.Add($nx); $pts.Add($ny)" & vbLf
    s = s & "                        if ($j + 2 -ge $arity) { $script:px = $nx; $script:py = $ny }" & vbLf
    s = s & "                    }" & vbLf
    s = s & "                }" & vbLf
    s = s & "            }" & vbLf
    s = s & "            $pts.Add($script:px); $pts.Add($script:py)" & vbLf
    s = s & "            $k += $arity" & vbLf
    s = s & "            if ($C -eq 'M') { $C = 'L' }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    $script:px = 0.0; $script:py = 0.0" & vbLf
    s = s & "    foreach ($t in $toks) {" & vbLf
    s = s & "        $v = $t.Value" & vbLf
    s = s & "        if ($v -cmatch '^[A-Za-z]$') {" & vbLf
    s = s & "            if ($v -eq 'F' -or $v -eq 'f') { $cmd = 'F'; $nums.Clear(); continue }" & vbLf
    s = s & "            if ($cmd -ne 'F') { & $flush $cmd $nums }" & vbLf
    s = s & "            $nums.Clear()" & vbLf
    s = s & "            $cmd = $v" & vbLf
    s = s & "        }" & vbLf
    s = s & "        else { $nums.Add((Num $v)) }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    if ($cmd -ne 'F') { & $flush $cmd $nums }" & vbLf
    s = s & "    return , $pts.ToArray()" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function GeometryFromElement($g) {" & vbLf
    s = s & "    # PathGeometry element -> @{ d = svg path; rule = 'nonzero'|'evenodd' }" & vbLf
    s = s & "    $rule = 'evenodd'" & vbLf
    s = s & "    if ($g.GetAttribute('FillRule') -eq 'NonZero') { $rule = 'nonzero' }" & vbLf
    s = s & "    $sb = New-Object System.Text.StringBuilder" & vbLf
    s = s & "    $fig = $g.GetAttribute('Figures')" & vbLf
    s = s & "    if ($fig) { [void]$sb.Append(($fig -replace '^\s*[Ff][01]\s*', '')) }" & vbLf
    s = s & "    foreach ($f in $g.ChildNodes) {" & vbLf
    s = s & "        if ($f.LocalName -ne 'PathFigure') { continue }" & vbLf
    s = s & "        [void]$sb.Append(' M ' + $f.GetAttribute('StartPoint'))" & vbLf
    s = s & "        foreach ($s in $f.ChildNodes) {" & vbLf
    s = s & "            switch ($s.LocalName) {" & vbLf
    s = s & "                'PolyLineSegment' { [void]$sb.Append(' L ' + $s.GetAttribute('Points')) }" & vbLf
    s = s & "                'PolyBezierSegment' { [void]$sb.Append(' C ' + $s.GetAttribute('Points')) }" & vbLf
    s = s & "                'PolyQuadraticBezierSegment' { [void]$sb.Append(' Q ' + $s.GetAttribute('Points')) }" & vbLf
    s = s & "                'ArcSegment' {" & vbLf
    s = s & "                    $sz = $s.GetAttribute('Size'); $rot = $s.GetAttribute('RotationAngle'); if (-not $rot) { $rot = '0' }" & vbLf
    s = s & "                    $large = $(if ($s.GetAttribute('IsLargeArc') -eq 'true') { '1' } else { '0' })" & vbLf
    s = s & "                    $sweep = $(if ($s.GetAttribute('SweepDirection') -eq 'Clockwise') { '1' } else { '0' })" & vbLf
    s = s & "                    [void]$sb.Append(' A ' + $sz + ' ' + $rot + ' ' + $large + ' ' + $sweep + ' ' + $s.GetAttribute('Point'))" & vbLf
    s = s & "                }" & vbLf
    s = s & "            }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        if ($f.GetAttribute('IsClosed') -eq 'true') { [void]$sb.Append(' Z') }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return @{ d = $sb.ToString().Trim(); rule = $rule }" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function Geometry([string]$attr, $node, [string]$propName) {" & vbLf
    s = s & "    # Data / Clip attribute (abbreviated syntax or StaticResource) or <X.Data> child element" & vbLf
    s = s & "    if ($attr) {" & vbLf
    s = s & "        if ($attr -match '^\{StaticResource\s+(.+)\}$') {" & vbLf
    s = s & "            $r = $Resources[$matches[1].Trim()]" & vbLf
    s = s & "            if ($r) { return (GeometryFromElement $r) }" & vbLf
    s = s & "            return $null" & vbLf
    s = s & "        }" & vbLf
    s = s & "        $rule = 'evenodd'" & vbLf
    s = s & "        if ($attr -match '^\s*[Ff]1') { $rule = 'nonzero' }" & vbLf
    s = s & "        return @{ d = ($attr -replace '^\s*[Ff][01]\s*', ''); rule = $rule }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    foreach ($c in $node.ChildNodes) {" & vbLf
    s = s & "        if ($c.LocalName -eq $propName) {" & vbLf
    s = s & "            foreach ($g in $c.ChildNodes) { if ($g.LocalName -eq 'PathGeometry') { return (GeometryFromElement $g) } }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return $null" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "# ------------------------------------------------------------------ colours and brushes" & vbLf
    s = s & "function Color([string]$s) {" & vbLf
    s = s & "    # -> @{ rgb = '#rrggbb'; a = 0..1 }" & vbLf
    s = s & "    $s = $s.Trim()" & vbLf
    s = s & "    if ($s -match '^#([0-9A-Fa-f]{8})$') {" & vbLf
    s = s & "        $h = $matches[1]" & vbLf
    s = s & "        return @{ rgb = '#' + $h.Substring(2, 6).ToLowerInvariant(); a = [Convert]::ToInt32($h.Substring(0, 2), 16) / 255.0 }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    if ($s -match '^#([0-9A-Fa-f]{6})$') { return @{ rgb = '#' + $matches[1].ToLowerInvariant(); a = 1.0 } }" & vbLf
    s = s & "    if ($s -match '^sc#') {" & vbLf
    s = s & "        $v = Numbers $s.Substring(3)" & vbLf
    s = s & "        if ($v.Count -eq 3) { $v = @(1.0) + $v }" & vbLf
    s = s & "        $c = @()" & vbLf
    s = s & "        for ($i = 1; $i -le 3; $i++) {" & vbLf
    s = s & "            $x = [Math]::Max(0.0, [Math]::Min(1.0, $v[$i]))" & vbLf
    s = s & "            $x = $(if ($x -le 0.0031308) { 12.92 * $x } else { 1.055 * [Math]::Pow($x, 1 / 2.4) - 0.055 })" & vbLf
    s = s & "            $c += [int][Math]::Round($x * 255)" & vbLf
    s = s & "        }" & vbLf
    s = s & "        return @{ rgb = ('#{0:x2}{1:x2}{2:x2}' -f $c[0], $c[1], $c[2]); a = [Math]::Max(0.0, [Math]::Min(1.0, $v[0])) }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return @{ rgb = '#000000'; a = 1.0 }" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function Stops($brush) {" & vbLf
    s = s & "    $sb = New-Object System.Text.StringBuilder" & vbLf
    s = s & "    foreach ($c in $brush.ChildNodes) {" & vbLf
    s = s & "        if ($c.LocalName -notlike '*.GradientStops') { continue }" & vbLf
    s = s & "        foreach ($s in $c.ChildNodes) {" & vbLf
    s = s & "            if ($s.LocalName -ne 'GradientStop') { continue }" & vbLf
    s = s & "            $col = Color $s.GetAttribute('Color')" & vbLf
    s = s & "            [void]$sb.Append('<stop offset=""' + (N (Num $s.GetAttribute('Offset'))) + '"" stop-color=""' + $col.rgb + '""')" & vbLf
    s = s & "            if ($col.a -lt 1) { [void]$sb.Append(' stop-opacity=""' + (N $col.a) + '""') }" & vbLf
    s = s & "            [void]$sb.Append('/>')" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return $sb.ToString()" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function BrushTransform($brush) {" & vbLf
    s = s & "    $t = $brush.GetAttribute('Transform')" & vbLf
    s = s & "    if ($t -match '^\{StaticResource') { $t = '' }" & vbLf
    s = s & "    if ($t) { return (ParseMatrix $t) }" & vbLf
    s = s & "    foreach ($c in $brush.ChildNodes) {" & vbLf
    s = s & "        if ($c.LocalName -like '*.Transform') {" & vbLf
    s = s & "            foreach ($m in $c.ChildNodes) { if ($m.LocalName -eq 'MatrixTransform') { return (ParseMatrix $m.GetAttribute('Matrix')) } }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return $null" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function ImageInfo([byte[]]$b) {" & vbLf
    s = s & "    # -> @{ mime; w; h; dpiX; dpiY } (pixel size and resolution)" & vbLf
    s = s & "    $info = @{ mime = 'image/png'; w = 0; h = 0; dpiX = 96.0; dpiY = 96.0 }" & vbLf
    s = s & "    if ($b.Length -gt 24 -and $b[0] -eq 0x89 -and $b[1] -eq 0x50) {" & vbLf
    s = s & "        $info.w = [int](U32 $b 16)" & vbLf
    s = s & "        $info.h = [int](U32 $b 20)" & vbLf
    s = s & "        $i = 8" & vbLf
    s = s & "        while ($i + 12 -le $b.Length) {" & vbLf
    s = s & "            $len = [int](U32 $b $i)" & vbLf
    s = s & "            $typ = [Text.Encoding]::ASCII.GetString($b, $i + 4, 4)" & vbLf
    s = s & "            if ($typ -eq 'pHYs' -and $b[$i + 16] -eq 1) {" & vbLf
    s = s & "                $px = U32 $b ($i + 8)" & vbLf
    s = s & "                $py = U32 $b ($i + 12)" & vbLf
    s = s & "                if ($px -gt 0) { $info.dpiX = $px * 0.0254 }" & vbLf
    s = s & "                if ($py -gt 0) { $info.dpiY = $py * 0.0254 }" & vbLf
    s = s & "            }" & vbLf
    s = s & "            if ($typ -eq 'IDAT' -or $typ -eq 'IEND' -or $len -lt 0) { break }" & vbLf
    s = s & "            $i += 12 + $len" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    elseif ($b.Length -gt 4 -and $b[0] -eq 0xFF -and $b[1] -eq 0xD8) {" & vbLf
    s = s & "        $info.mime = 'image/jpeg'" & vbLf
    s = s & "        $i = 2" & vbLf
    s = s & "        $lim = [Math]::Min($b.Length, 2097152)" & vbLf
    s = s & "        while ($i + 9 -lt $lim) {" & vbLf
    s = s & "            if ($b[$i] -ne 0xFF) { $i++; continue }" & vbLf
    s = s & "            $mk = $b[$i + 1]" & vbLf
    s = s & "            if ($mk -eq 0xDA -or $mk -eq 0xD9) { break }      # start of scan / end: no size marker found" & vbLf
    s = s & "            $len = ([int]$b[$i + 2] -shl 8) -bor $b[$i + 3]" & vbLf
    s = s & "            if ($mk -eq 0xE0 -and $b[$i + 4] -eq 0x4A -and $b[$i + 11] -eq 1) {" & vbLf
    s = s & "                $dx = ([int]$b[$i + 12] -shl 8) -bor $b[$i + 13]; $dy = ([int]$b[$i + 14] -shl 8) -bor $b[$i + 15]" & vbLf
    s = s & "                if ($dx -gt 0) { $info.dpiX = $dx; $info.dpiY = $dy }" & vbLf
    s = s & "            }" & vbLf
    s = s & "            if ($mk -ge 0xC0 -and $mk -le 0xCF -and $mk -ne 0xC4 -and $mk -ne 0xC8 -and $mk -ne 0xCC) {" & vbLf
    s = s & "                $info.h = ([int]$b[$i + 5] -shl 8) -bor $b[$i + 6]" & vbLf
    s = s & "                $info.w = ([int]$b[$i + 7] -shl 8) -bor $b[$i + 8]" & vbLf
    s = s & "                break" & vbLf
    s = s & "            }" & vbLf
    s = s & "            $i += 2 + $len" & vbLf
    HelperScript2 = s
End Function

Private Function HelperScript3() As String
    Dim s As String
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    elseif ($b.Length -gt 4 -and (($b[0] -eq 0x49 -and $b[1] -eq 0x49) -or ($b[0] -eq 0x4D -and $b[1] -eq 0x4D))) {" & vbLf
    s = s & "        $info.mime = 'image/tiff'" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return $info" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function Rect4([string]$s) {" & vbLf
    s = s & "    $v = Numbers $s" & vbLf
    s = s & "    if ($v.Count -lt 4) { return @(0.0, 0.0, 1.0, 1.0) }" & vbLf
    s = s & "    return @($v[0], $v[1], $v[2], $v[3])" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function Brush($brush, [string]$base) {" & vbLf
    s = s & "    # -> @{ paint = 'none' | '#rrggbb' | 'url(#id)'; a = opacity factor }" & vbLf
    s = s & "    if ($null -eq $brush) { return @{ paint = 'none'; a = 1.0 } }" & vbLf
    s = s & "    $op = 1.0" & vbLf
    s = s & "    if ($brush.GetAttribute('Opacity')) { $op = Num $brush.GetAttribute('Opacity') }" & vbLf
    s = s & "    switch ($brush.LocalName) {" & vbLf
    s = s & "        'SolidColorBrush' {" & vbLf
    s = s & "            $c = Color $brush.GetAttribute('Color')" & vbLf
    s = s & "            return @{ paint = $c.rgb; a = $c.a * $op }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        'LinearGradientBrush' {" & vbLf
    s = s & "            $id = NewId 'g'" & vbLf
    s = s & "            $s = Numbers $brush.GetAttribute('StartPoint'); $e = Numbers $brush.GetAttribute('EndPoint')" & vbLf
    s = s & "            $units = $(if ($brush.GetAttribute('MappingMode') -eq 'RelativeToBoundingBox') { 'objectBoundingBox' } else { 'userSpaceOnUse' })" & vbLf
    s = s & "            $spread = $brush.GetAttribute('SpreadMethod').ToLowerInvariant(); if (-not $spread) { $spread = 'pad' }" & vbLf
    s = s & "            [void]$Defs.Append('<linearGradient id=""' + $id + '"" gradientUnits=""' + $units + '"" spreadMethod=""' + $spread + '"" x1=""' + (N $s[0]) + '"" y1=""' + (N $s[1]) + '"" x2=""' + (N $e[0]) + '"" y2=""' + (N $e[1]) + '""')" & vbLf
    s = s & "            $m = BrushTransform $brush" & vbLf
    s = s & "            if ($m) { [void]$Defs.Append(' gradientTransform=""' + (MatText $m) + '""') }" & vbLf
    s = s & "            [void]$Defs.Append('>' + (Stops $brush) + '</linearGradient>')" & vbLf
    s = s & "            return @{ paint = 'url(#' + $id + ')'; a = $op }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        'RadialGradientBrush' {" & vbLf
    s = s & "            $id = NewId 'g'" & vbLf
    s = s & "            $c = Numbers $brush.GetAttribute('Center'); $o = Numbers $brush.GetAttribute('GradientOrigin')" & vbLf
    s = s & "            $rx = Num $brush.GetAttribute('RadiusX'); $ry = Num $brush.GetAttribute('RadiusY')" & vbLf
    s = s & "            if ($rx -le 0) { $rx = 0.001 }" & vbLf
    s = s & "            if ($ry -le 0) { $ry = 0.001 }" & vbLf
    s = s & "            $units = $(if ($brush.GetAttribute('MappingMode') -eq 'RelativeToBoundingBox') { 'objectBoundingBox' } else { 'userSpaceOnUse' })" & vbLf
    s = s & "            $spread = $brush.GetAttribute('SpreadMethod').ToLowerInvariant(); if (-not $spread) { $spread = 'pad' }" & vbLf
    s = s & "            $k = $ry / $rx" & vbLf
    s = s & "            $scale = @(1.0, 0.0, 0.0, $k, 0.0, ($c[1] - $k * $c[1]))" & vbLf
    s = s & "            $m = BrushTransform $brush" & vbLf
    s = s & "            if ($m) { $scale = MatMul $scale $m }" & vbLf
    s = s & "            $fy = $c[1] + ($o[1] - $c[1]) / $k" & vbLf
    s = s & "            [void]$Defs.Append('<radialGradient id=""' + $id + '"" gradientUnits=""' + $units + '"" spreadMethod=""' + $spread + '"" cx=""' + (N $c[0]) + '"" cy=""' + (N $c[1]) + '"" r=""' + (N $rx) + '"" fx=""' + (N $o[0]) + '"" fy=""' + (N $fy) + '"" gradientTransform=""' + (MatText $scale) + '"">' + (Stops $brush) + '</radialGradient>')" & vbLf
    s = s & "            return @{ paint = 'url(#' + $id + ')'; a = $op }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        'ImageBrush' {" & vbLf
    s = s & "            $src = $brush.GetAttribute('ImageSource')" & vbLf
    s = s & "            if ($src -match '^\{') { return @{ paint = 'none'; a = 1.0 } }" & vbLf
    s = s & "            $part = Resolve $base $src" & vbLf
    s = s & "            $bytes = ReadBytes $part" & vbLf
    s = s & "            if ($null -eq $bytes) { return @{ paint = 'none'; a = 1.0 } }" & vbLf
    s = s & "            $info = ImageInfo $bytes" & vbLf
    s = s & "            $vb = Rect4 $brush.GetAttribute('Viewbox'); $vp = Rect4 $brush.GetAttribute('Viewport')" & vbLf
    s = s & "            $nw = $(if ($info.w -gt 0) { $info.w * 96.0 / $info.dpiX } else { $vb[2] })" & vbLf
    s = s & "            $nh = $(if ($info.h -gt 0) { $info.h * 96.0 / $info.dpiY } else { $vb[3] })" & vbLf
    s = s & "            if ($brush.GetAttribute('ViewboxUnits') -eq 'RelativeToBoundingBox') {" & vbLf
    s = s & "                $vb = @(($vb[0] * $nw), ($vb[1] * $nh), ($vb[2] * $nw), ($vb[3] * $nh))" & vbLf
    s = s & "            }" & vbLf
    s = s & "            if ($vb[2] -le 0) { $vb[2] = 1 }" & vbLf
    s = s & "            if ($vb[3] -le 0) { $vb[3] = 1 }" & vbLf
    s = s & "            $sx = $vp[2] / $vb[2]; $sy = $vp[3] / $vb[3]" & vbLf
    s = s & "            $id = NewId 'p'" & vbLf
    s = s & "            [void]$Defs.Append('<pattern id=""' + $id + '"" patternUnits=""userSpaceOnUse"" x=""' + (N $vp[0]) + '"" y=""' + (N $vp[1]) + '"" width=""' + (N $vp[2]) + '"" height=""' + (N $vp[3]) + '""')" & vbLf
    s = s & "            $m = BrushTransform $brush" & vbLf
    s = s & "            if ($m) { [void]$Defs.Append(' patternTransform=""' + (MatText $m) + '""') }" & vbLf
    s = s & "            [void]$Defs.Append('><image x=""' + (N (-$vb[0] * $sx)) + '"" y=""' + (N (-$vb[1] * $sy)) + '"" width=""' + (N ($nw * $sx)) + '"" height=""' + (N ($nh * $sy)) + '"" preserveAspectRatio=""none"" xlink:href=""data:' + $info.mime + ';base64,' + [Convert]::ToBase64String($bytes) + '""/></pattern>')" & vbLf
    s = s & "            return @{ paint = 'url(#' + $id + ')'; a = $op }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return @{ paint = 'none'; a = 1.0 }" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function BrushOf($node, [string]$attr, [string]$prop, [string]$base) {" & vbLf
    s = s & "    $v = $node.GetAttribute($attr)" & vbLf
    s = s & "    if ($v) {" & vbLf
    s = s & "        if ($v -match '^\{StaticResource\s+(.+)\}$') { return (Brush $Resources[$matches[1].Trim()] $base) }" & vbLf
    s = s & "        $c = Color $v" & vbLf
    s = s & "        return @{ paint = $c.rgb; a = $c.a }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    foreach ($c in $node.ChildNodes) {" & vbLf
    s = s & "        if ($c.LocalName -eq $prop) {" & vbLf
    s = s & "            foreach ($b in $c.ChildNodes) { if ($b.NodeType -eq 'Element') { return (Brush $b $base) } }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return $null" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "# ------------------------------------------------------------------ fonts" & vbLf
    s = s & "function U16([byte[]]$b, [int]$o) { return ([int]$b[$o] -shl 8) -bor $b[$o + 1] }" & vbLf
    s = s & "function U32([byte[]]$b, [int]$o) { return ([int64]$b[$o] -shl 24) -bor ([int64]$b[$o + 1] -shl 16) -bor ([int64]$b[$o + 2] -shl 8) -bor [int64]$b[$o + 3] }" & vbLf
    s = s & "function S16([byte[]]$b, [int]$o) { $v = U16 $b $o; if ($v -ge 32768) { $v -= 65536 }; return $v }" & vbLf
    s = s & "" & vbLf
    s = s & "function FontTables([byte[]]$b) {" & vbLf
    s = s & "    $t = @{}" & vbLf
    s = s & "    if ($b.Length -lt 12) { return $t }" & vbLf
    s = s & "    $off = 0" & vbLf
    s = s & "    if ([Text.Encoding]::ASCII.GetString($b, 0, 4) -eq 'ttcf') { $off = [int](U32 $b 12) }" & vbLf
    s = s & "    $n = U16 $b ($off + 4)" & vbLf
    s = s & "    for ($i = 0; $i -lt $n; $i++) {" & vbLf
    s = s & "        $r = $off + 12 + 16 * $i" & vbLf
    s = s & "        if ($r + 16 -gt $b.Length) { break }" & vbLf
    s = s & "        $t[[Text.Encoding]::ASCII.GetString($b, $r, 4)] = @([int](U32 $b ($r + 8)), [int](U32 $b ($r + 12)))" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return $t" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function IsFont([byte[]]$b) {" & vbLf
    s = s & "    if ($b.Length -lt 4) { return $false }" & vbLf
    s = s & "    $tag = [Text.Encoding]::ASCII.GetString($b, 0, 4)" & vbLf
    s = s & "    return ($tag -eq 'OTTO' -or $tag -eq 'true' -or $tag -eq 'ttcf' -or ($b[0] -eq 0 -and $b[1] -eq 1 -and $b[2] -eq 0 -and $b[3] -eq 0))" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function Deobfuscate([byte[]]$b, [string]$part) {" & vbLf
    s = s & "    if (IsFont $b) { return $b }" & vbLf
    s = s & "    $m = [regex]::Match($part, '([0-9A-Fa-f]{8})-?([0-9A-Fa-f]{4})-?([0-9A-Fa-f]{4})-?([0-9A-Fa-f]{4})-?([0-9A-Fa-f]{12})')" & vbLf
    s = s & "    if (-not $m.Success) { return $b }" & vbLf
    s = s & "    $hex = ($m.Groups[1].Value + $m.Groups[2].Value + $m.Groups[3].Value + $m.Groups[4].Value + $m.Groups[5].Value)" & vbLf
    s = s & "    $fwd = New-Object byte[] 16" & vbLf
    s = s & "    for ($i = 0; $i -lt 16; $i++) { $fwd[$i] = [Convert]::ToByte($hex.Substring(2 * $i, 2), 16) }" & vbLf
    s = s & "    $rev = New-Object byte[] 16" & vbLf
    s = s & "    for ($i = 0; $i -lt 16; $i++) { $rev[$i] = $fwd[15 - $i] }" & vbLf
    s = s & "    foreach ($key in @($rev, $fwd)) {" & vbLf
    s = s & "        $c = [byte[]]$b.Clone()" & vbLf
    s = s & "        for ($i = 0; $i -lt 32 -and $i -lt $c.Length; $i++) { $c[$i] = $c[$i] -bxor $key[$i % 16] }" & vbLf
    s = s & "        if (IsFont $c) { return $c }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return $b" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function FontInfo([string]$part) {" & vbLf
    s = s & "    if ($Fonts.ContainsKey($part)) { return $Fonts[$part] }" & vbLf
    s = s & "    $info = @{ family = 'sans-serif'; local = ''; weight = 400; italic = $false; upm = 1000; adv = $null; cmap = $null; rev = $null; id = ''; data = $null; bytes = $null; tables = $null; metrics = $false; mapped = $false }" & vbLf
    s = s & "    $Fonts[$part] = $info" & vbLf
    s = s & "    $b = ReadBytes $part" & vbLf
    s = s & "    if ($null -eq $b) { return $info }" & vbLf
    s = s & "    $b = Deobfuscate $b $part" & vbLf
    s = s & "    if (-not (IsFont $b)) { return $info }" & vbLf
    s = s & "    $info.bytes = $b" & vbLf
    s = s & "    $t = FontTables $b" & vbLf
    s = s & "    $info.tables = $t" & vbLf
    s = s & "    try {" & vbLf
    s = s & "        if ($t.ContainsKey('name')) {" & vbLf
    s = s & "            $o = $t['name'][0]" & vbLf
    s = s & "            $count = U16 $b ($o + 2); $str = $o + (U16 $b ($o + 4))" & vbLf
    s = s & "            $best = @{}" & vbLf
    s = s & "            for ($i = 0; $i -lt $count -and $i -lt 400; $i++) {" & vbLf
    s = s & "                $r = $o + 6 + 12 * $i" & vbLf
    s = s & "                $plat = U16 $b $r; $lang = U16 $b ($r + 4); $nid = U16 $b ($r + 6)" & vbLf
    s = s & "                $len = U16 $b ($r + 8); $off = U16 $b ($r + 10)" & vbLf
    s = s & "                if ($plat -ne 3 -or ($nid -ne 1 -and $nid -ne 16)) { continue }" & vbLf
    s = s & "                $s = [Text.Encoding]::BigEndianUnicode.GetString($b, $str + $off, $len)" & vbLf
    s = s & "                $key = $(if ($lang -eq 0x409) { 'en' } elseif ($lang -eq 0x804 -or $lang -eq 0x404 -or $lang -eq 0x411) { 'zh' } else { 'x' }) + $nid" & vbLf
    s = s & "                $best[$key] = $s" & vbLf
    s = s & "            }" & vbLf
    s = s & "            foreach ($k in @('en16', 'en1', 'x16', 'x1')) { if ($best[$k]) { $info.family = $best[$k]; break } }" & vbLf
    s = s & "            foreach ($k in @('zh16', 'zh1')) { if ($best[$k] -and $best[$k] -ne $info.family) { $info.local = $best[$k]; break } }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        if ($t.ContainsKey('OS/2')) {" & vbLf
    s = s & "            $o = $t['OS/2'][0]" & vbLf
    s = s & "            $info.weight = U16 $b ($o + 4)" & vbLf
    s = s & "            $info.italic = ((U16 $b ($o + 62)) -band 1) -eq 1" & vbLf
    s = s & "        }" & vbLf
    s = s & "        if ($t.ContainsKey('head')) { $info.upm = U16 $b ($t['head'][0] + 18) }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    catch { }" & vbLf
    s = s & "    $info.id = NewId 'f'" & vbLf
    s = s & "    # Large fonts (a complete CJK font, for example) are referenced by name only: embedding them" & vbLf
    s = s & "    # would make the SVG huge and slow to write. TrueType collections cannot be embedded." & vbLf
    s = s & "    if (-not $NoEmbedFonts -and $b.Length -le $MaxFontBytes -and [Text.Encoding]::ASCII.GetString($b, 0, 4) -ne 'ttcf') {" & vbLf
    s = s & "        $info.data = $b" & vbLf
    s = s & "        $script:FontsEmbedded++" & vbLf
    s = s & "        Report 88 'font' $script:FontsEmbedded 0" & vbLf
    s = s & "        [void]$Defs.Append('<style>@font-face{font-family:''' + $info.id + ''';src:url(data:font/ttf;base64,' + [Convert]::ToBase64String($b) + ');}</style>')" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return $info" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "# Advance widths (hmtx) - only read when a glyph run has no advances of its own." & vbLf
    s = s & "function EnsureMetrics($info) {" & vbLf
    s = s & "    if ($info.metrics) { return }" & vbLf
    s = s & "    $info.metrics = $true" & vbLf
    s = s & "    $t = $info.tables; $b = $info.bytes" & vbLf
    s = s & "    if ($null -eq $t -or $null -eq $b) { return }" & vbLf
    s = s & "    try {" & vbLf
    s = s & "        if ($t.ContainsKey('hhea') -and $t.ContainsKey('hmtx')) {" & vbLf
    s = s & "            $nh = U16 $b ($t['hhea'][0] + 34)" & vbLf
    s = s & "            $o = $t['hmtx'][0]" & vbLf
    s = s & "            $adv = New-Object int[] ([Math]::Max($nh, 1))" & vbLf
    s = s & "            for ($i = 0; $i -lt $nh; $i++) { $adv[$i] = ([int]$b[$o + 4 * $i] -shl 8) -bor $b[$o + 4 * $i + 1] }" & vbLf
    s = s & "            $info.adv = $adv" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    catch { }" & vbLf
    s = s & "}" & vbLf
    HelperScript3 = s
End Function

Private Function HelperScript4() As String
    Dim s As String
    s = s & "" & vbLf
    s = s & "# Character map - only read when a glyph run lacks glyph indices or characters." & vbLf
    s = s & "function EnsureCmap($info) {" & vbLf
    s = s & "    if ($info.mapped) { return }" & vbLf
    s = s & "    $info.mapped = $true" & vbLf
    s = s & "    $t = $info.tables; $b = $info.bytes" & vbLf
    s = s & "    if ($null -eq $t -or $null -eq $b -or -not $t.ContainsKey('cmap')) { return }" & vbLf
    s = s & "    try {" & vbLf
    s = s & "        $o = $t['cmap'][0]" & vbLf
    s = s & "        $n = U16 $b ($o + 2)" & vbLf
    s = s & "        $sub = -1" & vbLf
    s = s & "        for ($i = 0; $i -lt $n; $i++) {" & vbLf
    s = s & "            $r = $o + 4 + 8 * $i" & vbLf
    s = s & "            $plat = U16 $b $r; $eid = U16 $b ($r + 2); $so = $o + [int](U32 $b ($r + 4))" & vbLf
    s = s & "            $fmt = U16 $b $so" & vbLf
    s = s & "            if ($plat -eq 3 -and ($eid -eq 1 -or $eid -eq 10 -or $eid -eq 0) -and ($fmt -eq 4 -or $fmt -eq 12)) { $sub = $so; if ($fmt -eq 12) { break } }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        if ($sub -lt 0) { return }" & vbLf
    s = s & "        $map = New-Object 'System.Collections.Generic.Dictionary[int,int]'" & vbLf
    s = s & "        $fmt = U16 $b $sub" & vbLf
    s = s & "        if ($fmt -eq 4) {" & vbLf
    s = s & "            $segX2 = U16 $b ($sub + 6)" & vbLf
    s = s & "            $ends = $sub + 14; $starts = $ends + $segX2 + 2; $deltas = $starts + $segX2; $ranges = $deltas + $segX2" & vbLf
    s = s & "            for ($s = 0; $s -lt $segX2 / 2; $s++) {" & vbLf
    s = s & "                $e = U16 $b ($ends + 2 * $s); $st = U16 $b ($starts + 2 * $s)" & vbLf
    s = s & "                $dl = U16 $b ($deltas + 2 * $s); $ro = U16 $b ($ranges + 2 * $s)" & vbLf
    s = s & "                if ($st -eq 0xFFFF) { continue }" & vbLf
    s = s & "                for ($c = $st; $c -le $e; $c++) {" & vbLf
    s = s & "                    if ($ro -eq 0) { $g = ($c + $dl) % 65536 }" & vbLf
    s = s & "                    else {" & vbLf
    s = s & "                        $ga = $ranges + 2 * $s + $ro + 2 * ($c - $st)" & vbLf
    s = s & "                        $g = ([int]$b[$ga] -shl 8) -bor $b[$ga + 1]" & vbLf
    s = s & "                        if ($g -ne 0) { $g = ($g + $dl) % 65536 }" & vbLf
    s = s & "                    }" & vbLf
    s = s & "                    if ($g -ne 0) { $map[$c] = $g }" & vbLf
    s = s & "                }" & vbLf
    s = s & "                CheckTime" & vbLf
    s = s & "            }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        else {" & vbLf
    s = s & "            $ng = [int](U32 $b ($sub + 12))" & vbLf
    s = s & "            for ($s = 0; $s -lt $ng; $s++) {" & vbLf
    s = s & "                $r = $sub + 16 + 12 * $s" & vbLf
    s = s & "                $st = [int](U32 $b $r); $e = [int](U32 $b ($r + 4)); $g0 = [int](U32 $b ($r + 8))" & vbLf
    s = s & "                if ($e - $st -gt 70000) { continue }" & vbLf
    s = s & "                for ($c = $st; $c -le $e; $c++) { $map[$c] = $g0 + ($c - $st) }" & vbLf
    s = s & "                CheckTime" & vbLf
    s = s & "            }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        $info.cmap = $map" & vbLf
    s = s & "    }" & vbLf
    s = s & "    catch { }" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function GlyphAdvance($info, [int]$gid) {" & vbLf
    s = s & "    EnsureMetrics $info" & vbLf
    s = s & "    if ($null -eq $info.adv -or $info.adv.Count -eq 0) { return 0.0 }" & vbLf
    s = s & "    if ($gid -ge $info.adv.Count) { $gid = $info.adv.Count - 1 }" & vbLf
    s = s & "    if ($gid -lt 0) { return 0.0 }" & vbLf
    s = s & "    return $info.adv[$gid] / [double]$info.upm" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function GidOf($info, [int]$cp) {" & vbLf
    s = s & "    EnsureCmap $info" & vbLf
    s = s & "    if ($null -eq $info.cmap) { return -1 }" & vbLf
    s = s & "    $g = 0" & vbLf
    s = s & "    if ($info.cmap.TryGetValue($cp, [ref]$g)) { return [int]$g }" & vbLf
    s = s & "    return -1" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function CharFromGid($info, [int]$gid) {" & vbLf
    s = s & "    EnsureCmap $info" & vbLf
    s = s & "    if ($null -eq $info.cmap) { return $null }" & vbLf
    s = s & "    if ($null -eq $info.rev) {" & vbLf
    s = s & "        $info.rev = @{}" & vbLf
    s = s & "        foreach ($k in $info.cmap.Keys) { if (-not $info.rev.ContainsKey($info.cmap[$k])) { $info.rev[$info.cmap[$k]] = $k } }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    $cp = $info.rev[$gid]" & vbLf
    s = s & "    if ($null -eq $cp) { return $null }" & vbLf
    s = s & "    return [char]::ConvertFromUtf32([int]$cp)" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "# ------------------------------------------------------------------ elements" & vbLf
    s = s & "function TransformOf($node, [string]$prop) {" & vbLf
    s = s & "    $t = $node.GetAttribute('RenderTransform')" & vbLf
    s = s & "    if ($t -match '^\{StaticResource\s+(.+)\}$') {" & vbLf
    s = s & "        $r = $Resources[$matches[1].Trim()]" & vbLf
    s = s & "        if ($r) { return (ParseMatrix $r.GetAttribute('Matrix')) }" & vbLf
    s = s & "        return $null" & vbLf
    s = s & "    }" & vbLf
    s = s & "    if ($t) { return (ParseMatrix $t) }" & vbLf
    s = s & "    foreach ($c in $node.ChildNodes) {" & vbLf
    s = s & "        if ($c.LocalName -eq $prop) {" & vbLf
    s = s & "            foreach ($m in $c.ChildNodes) { if ($m.LocalName -eq 'MatrixTransform') { return (ParseMatrix $m.GetAttribute('Matrix')) } }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return $null" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function CommonAttrs($node, [string]$kind, $ctm, [ref]$mOut, [string]$base) {" & vbLf
    s = s & "    # transform, clip and opacity of an element -> SVG attribute string; mOut receives the new CTM" & vbLf
    s = s & "    $s = ''" & vbLf
    s = s & "    $m = TransformOf $node ($kind + '.RenderTransform')" & vbLf
    s = s & "    $cur = $ctm" & vbLf
    s = s & "    if ($m) { $s += ' transform=""' + (MatText $m) + '""'; $cur = MatMul $m $ctm }" & vbLf
    s = s & "    $mOut.Value = $cur" & vbLf
    s = s & "    $clip = Geometry $node.GetAttribute('Clip') $node ($kind + '.Clip')" & vbLf
    s = s & "    if ($clip -and $clip.d) {" & vbLf
    s = s & "        $id = NewId 'c'" & vbLf
    s = s & "        [void]$Defs.Append('<clipPath id=""' + $id + '""><path d=""' + $clip.d + '"" clip-rule=""' + $clip.rule + '""/></clipPath>')" & vbLf
    s = s & "        $s += ' clip-path=""url(#' + $id + ')""'" & vbLf
    s = s & "    }" & vbLf
    s = s & "    $op = 1.0" & vbLf
    s = s & "    if ($node.GetAttribute('Opacity')) { $op = Num $node.GetAttribute('Opacity') }" & vbLf
    s = s & "    $mask = BrushOf $node 'OpacityMask' ($kind + '.OpacityMask') $base" & vbLf
    s = s & "    if ($mask -and $mask.paint -match '^#') { $op *= $mask.a }" & vbLf
    s = s & "    if ($op -lt 0.999) { $s += ' opacity=""' + (N $op) + '""' }" & vbLf
    s = s & "    return $s" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function LoadResources($node, [string]$base) {" & vbLf
    s = s & "    foreach ($c in $node.ChildNodes) {" & vbLf
    s = s & "        if ($c.LocalName -notlike '*.Resources') { continue }" & vbLf
    s = s & "        foreach ($d in $c.ChildNodes) {" & vbLf
    s = s & "            if ($d.LocalName -ne 'ResourceDictionary') { continue }" & vbLf
    s = s & "            $src = $d.GetAttribute('Source')" & vbLf
    s = s & "            if ($src) {" & vbLf
    s = s & "                $part = Resolve $base $src" & vbLf
    s = s & "                $x = ReadXml $part" & vbLf
    s = s & "                if ($x) { foreach ($r in $x.DocumentElement.ChildNodes) { if ($r.NodeType -eq 'Element') { $k = $r.GetAttribute('Key', $KeyNs); if ($k) { $Resources[$k] = $r } } } }" & vbLf
    s = s & "            }" & vbLf
    s = s & "            foreach ($r in $d.ChildNodes) {" & vbLf
    s = s & "                if ($r.NodeType -ne 'Element') { continue }" & vbLf
    s = s & "                $k = $r.GetAttribute('Key', $KeyNs)" & vbLf
    s = s & "                if ($k) { $Resources[$k] = $r }" & vbLf
    s = s & "            }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function IsPageBackground($pts, $ctm) {" & vbLf
    s = s & "    if ($null -eq $pts -or $pts.Count -lt 4) { return $false }" & vbLf
    s = s & "    $x0 = [double]::MaxValue; $y0 = [double]::MaxValue; $x1 = -[double]::MaxValue; $y1 = -[double]::MaxValue" & vbLf
    s = s & "    for ($i = 0; $i + 1 -lt $pts.Count; $i += 2) {" & vbLf
    s = s & "        $x = $ctm[0] * $pts[$i] + $ctm[2] * $pts[$i + 1] + $ctm[4]" & vbLf
    s = s & "        $y = $ctm[1] * $pts[$i] + $ctm[3] * $pts[$i + 1] + $ctm[5]" & vbLf
    s = s & "        $x0 = [Math]::Min($x0, $x); $y0 = [Math]::Min($y0, $y); $x1 = [Math]::Max($x1, $x); $y1 = [Math]::Max($y1, $y)" & vbLf
    s = s & "    }" & vbLf
    s = s & "    return ($x0 -le 0.5 -and $y0 -le 0.5 -and $x1 -ge $script:PageW - 0.5 -and $y1 -ge $script:PageH - 0.5)" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function DoPath($node, $ctm, [string]$base, [System.Text.StringBuilder]$sb) {" & vbLf
    s = s & "    Tick" & vbLf
    s = s & "    $g = Geometry $node.GetAttribute('Data') $node 'Path.Data'" & vbLf
    s = s & "    if ($null -eq $g -or -not $g.d) { return }" & vbLf
    s = s & "    $pts = FigureBounds $g.d" & vbLf
    s = s & "    $m = $null" & vbLf
    s = s & "    $attrs = CommonAttrs $node 'Path' $ctm ([ref]$m) $base" & vbLf
    s = s & "    if (IsPageBackground $pts $m) { return }" & vbLf
    s = s & "    $fill = BrushOf $node 'Fill' 'Path.Fill' $base" & vbLf
    s = s & "    $stroke = BrushOf $node 'Stroke' 'Path.Stroke' $base" & vbLf
    s = s & "    [void]$sb.Append('<path d=""' + $g.d + '""' + $attrs)" & vbLf
    s = s & "    if ($null -eq $fill -or $fill.paint -eq 'none') { [void]$sb.Append(' fill=""none""') }" & vbLf
    s = s & "    else {" & vbLf
    s = s & "        [void]$sb.Append(' fill=""' + $fill.paint + '""')" & vbLf
    s = s & "        if ($fill.a -lt 0.999) { [void]$sb.Append(' fill-opacity=""' + (N $fill.a) + '""') }" & vbLf
    s = s & "        if ($g.rule -eq 'evenodd') { [void]$sb.Append(' fill-rule=""evenodd""') }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    $pad = 0.0" & vbLf
    s = s & "    if ($stroke -and $stroke.paint -ne 'none') {" & vbLf
    s = s & "        $w = 1.0" & vbLf
    s = s & "        if ($node.GetAttribute('StrokeThickness')) { $w = Num $node.GetAttribute('StrokeThickness') }" & vbLf
    s = s & "        $pad = $w / 2" & vbLf
    s = s & "        [void]$sb.Append(' stroke=""' + $stroke.paint + '"" stroke-width=""' + (N $w) + '""')" & vbLf
    s = s & "        if ($stroke.a -lt 0.999) { [void]$sb.Append(' stroke-opacity=""' + (N $stroke.a) + '""') }" & vbLf
    s = s & "        $join = $node.GetAttribute('StrokeLineJoin')" & vbLf
    s = s & "        if ($join -eq 'Round') { [void]$sb.Append(' stroke-linejoin=""round""') }" & vbLf
    s = s & "        elseif ($join -eq 'Bevel') { [void]$sb.Append(' stroke-linejoin=""bevel""') }" & vbLf
    s = s & "        else {" & vbLf
    s = s & "            $ml = 10.0" & vbLf
    s = s & "            if ($node.GetAttribute('StrokeMiterLimit')) { $ml = Num $node.GetAttribute('StrokeMiterLimit') }" & vbLf
    s = s & "            [void]$sb.Append(' stroke-miterlimit=""' + (N ([Math]::Max($ml, 1.0))) + '""')" & vbLf
    s = s & "        }" & vbLf
    s = s & "        $cap = $node.GetAttribute('StrokeStartLineCap')" & vbLf
    s = s & "        if ($cap -eq 'Round') { [void]$sb.Append(' stroke-linecap=""round""') }" & vbLf
    s = s & "        elseif ($cap -eq 'Square') { [void]$sb.Append(' stroke-linecap=""square""') }" & vbLf
    s = s & "        $dash = $node.GetAttribute('StrokeDashArray')" & vbLf
    s = s & "        if ($dash) {" & vbLf
    s = s & "            $dv = Numbers $dash" & vbLf
    s = s & "            if ($dv.Count -gt 0) {" & vbLf
    s = s & "                [void]$sb.Append(' stroke-dasharray=""' + (($dv | ForEach-Object { N ([Math]::Max($_ * $w, 0.01)) }) -join ' ') + '""')" & vbLf
    s = s & "                if ($node.GetAttribute('StrokeDashOffset')) { [void]$sb.Append(' stroke-dashoffset=""' + (N ((Num $node.GetAttribute('StrokeDashOffset')) * $w)) + '""') }" & vbLf
    s = s & "                if ($node.GetAttribute('StrokeDashCap') -eq 'Round') { [void]$sb.Append(' stroke-linecap=""round""') }" & vbLf
    s = s & "            }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    [void]$sb.Append('/>')" & vbLf
    s = s & "    if (($fill -and $fill.paint -ne 'none') -or ($stroke -and $stroke.paint -ne 'none')) { AddBox $m $pts $pad }" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    HelperScript4 = s
End Function

Private Function HelperScript5() As String
    Dim s As String
    s = s & "function DoGlyphs($node, $ctm, [string]$base, [System.Text.StringBuilder]$sb) {" & vbLf
    s = s & "    Tick" & vbLf
    s = s & "    $fontPart = Resolve $base $node.GetAttribute('FontUri')" & vbLf
    s = s & "    $info = FontInfo $fontPart" & vbLf
    s = s & "    $em = Num $node.GetAttribute('FontRenderingEmSize')" & vbLf
    s = s & "    $ox = Num $node.GetAttribute('OriginX'); $oy = Num $node.GetAttribute('OriginY')" & vbLf
    s = s & "    $text = $node.GetAttribute('UnicodeString')" & vbLf
    s = s & "    if ($text.StartsWith('{}')) { $text = $text.Substring(2) }" & vbLf
    s = s & "    $idx = $node.GetAttribute('Indices')" & vbLf
    s = s & "    $entries = @()" & vbLf
    s = s & "    if ($idx) { $entries = $idx.Split(';') }" & vbLf
    s = s & "    $hasClusters = $idx -match '\('" & vbLf
    s = s & "    $rtl = $false" & vbLf
    s = s & "    if ($node.GetAttribute('BidiLevel')) { $rtl = ((Num $node.GetAttribute('BidiLevel')) % 2) -eq 1 }" & vbLf
    s = s & "    # code points of the text" & vbLf
    s = s & "    $cps = New-Object System.Collections.Generic.List[int]" & vbLf
    s = s & "    for ($i = 0; $i -lt $text.Length; $i++) {" & vbLf
    s = s & "        if ([char]::IsHighSurrogate($text[$i]) -and $i + 1 -lt $text.Length) { $cps.Add([char]::ConvertToUtf32($text[$i], $text[$i + 1])); $i++ }" & vbLf
    s = s & "        else { $cps.Add([int]$text[$i]) }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    if ($cps.Count -eq 0 -and $entries.Count -gt 0) {" & vbLf
    s = s & "        # glyph indices only: recover the characters from the font's cmap" & vbLf
    s = s & "        $chars = ''" & vbLf
    s = s & "        foreach ($e in $entries) {" & vbLf
    s = s & "            $gi = ($e.Split(',')[0] -replace '\(.*\)', '').Trim()" & vbLf
    s = s & "            if ($gi -match '^\d+$') { $c = CharFromGid $info ([int]$gi); if ($c) { $chars += $c } }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        if (-not $chars) { return }" & vbLf
    s = s & "        $text = $chars" & vbLf
    s = s & "        for ($i = 0; $i -lt $text.Length; $i++) {" & vbLf
    s = s & "            if ([char]::IsHighSurrogate($text[$i]) -and $i + 1 -lt $text.Length) { $cps.Add([char]::ConvertToUtf32($text[$i], $text[$i + 1])); $i++ }" & vbLf
    s = s & "            else { $cps.Add([int]$text[$i]) }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    if ($cps.Count -eq 0) { return }" & vbLf
    s = s & "    # x position of every character from the advance widths" & vbLf
    s = s & "    $xs = New-Object System.Collections.Generic.List[string]" & vbLf
    s = s & "    $x = $ox" & vbLf
    s = s & "    $dir = $(if ($rtl) { -1.0 } else { 1.0 })" & vbLf
    s = s & "    $total = 0.0" & vbLf
    s = s & "    for ($i = 0; $i -lt $cps.Count; $i++) {" & vbLf
    s = s & "        $gid = -1; $adv = $null; $uoff = 0.0" & vbLf
    s = s & "        if (-not $hasClusters -and $i -lt $entries.Count) {" & vbLf
    s = s & "            $f = $entries[$i].Split(',')" & vbLf
    s = s & "            if ($f[0].Trim() -match '^\d+$') { $gid = [int]$f[0].Trim() }" & vbLf
    s = s & "            if ($f.Count -gt 1 -and $f[1].Trim()) { $adv = (Num $f[1]) / 100.0 }" & vbLf
    s = s & "            if ($f.Count -gt 2 -and $f[2].Trim()) { $uoff = (Num $f[2]) / 100.0 * $em }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        if ($gid -lt 0) { $gid = GidOf $info $cps[$i] }" & vbLf
    s = s & "        if ($null -eq $adv) { $adv = GlyphAdvance $info $gid }" & vbLf
    s = s & "        $xs.Add((N ($x + $dir * $uoff)))" & vbLf
    s = s & "        $x += $dir * $adv * $em" & vbLf
    s = s & "        $total += $adv * $em" & vbLf
    s = s & "    }" & vbLf
    s = s & "    $fill = BrushOf $node 'Fill' 'Glyphs.Fill' $base" & vbLf
    s = s & "    if ($null -eq $fill) { $fill = @{ paint = '#000000'; a = 1.0 } }" & vbLf
    s = s & "    $m = $null" & vbLf
    s = s & "    $attrs = CommonAttrs $node 'Glyphs' $ctm ([ref]$m) $base" & vbLf
    s = s & "    $fam = ""'"" + $info.family.Replace(""'"", '') + ""'""" & vbLf
    s = s & "    if ($info.local) { $fam += "", '"" + $info.local.Replace(""'"", '') + ""'"" }" & vbLf
    s = s & "    if ($info.data -and -not $NoEmbedFonts) { $fam = $info.id + ', ' + $fam }" & vbLf
    s = s & "    $weight = $info.weight" & vbLf
    s = s & "    $sim = $node.GetAttribute('StyleSimulations')" & vbLf
    s = s & "    if ($sim -match 'Bold') { $weight = 700 }" & vbLf
    s = s & "    [void]$sb.Append('<text xml:space=""preserve""' + $attrs + ' x=""' + ($xs -join ' ') + '"" y=""' + (N $oy) + '"" font-family=""' + (Esc $fam) + '"" font-size=""' + (N $em) + '""')" & vbLf
    s = s & "    if ($weight -ge 600) { [void]$sb.Append(' font-weight=""' + $weight + '""') }" & vbLf
    s = s & "    if ($info.italic -or $sim -match 'Italic') { [void]$sb.Append(' font-style=""italic""') }" & vbLf
    s = s & "    if ($fill.paint -ne '#000000') { [void]$sb.Append(' fill=""' + $fill.paint + '""') }" & vbLf
    s = s & "    if ($fill.a -lt 0.999) { [void]$sb.Append(' fill-opacity=""' + (N $fill.a) + '""') }" & vbLf
    s = s & "    [void]$sb.Append('>' + (Esc $text) + '</text>')" & vbLf
    s = s & "    $x0 = [Math]::Min($ox, $ox + $dir * $total); $x1 = [Math]::Max($ox, $ox + $dir * $total)" & vbLf
    s = s & "    AddBox $m @($x0, ($oy - 0.95 * $em), $x1, ($oy + 0.3 * $em)) 0.0" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function DoChildren($node, $ctm, [string]$base, [System.Text.StringBuilder]$sb) {" & vbLf
    s = s & "    foreach ($c in $node.ChildNodes) {" & vbLf
    s = s & "        if ($c.NodeType -ne 'Element') { continue }" & vbLf
    s = s & "        switch ($c.LocalName) {" & vbLf
    s = s & "            'Path' { DoPath $c $ctm $base $sb }" & vbLf
    s = s & "            'Glyphs' { DoGlyphs $c $ctm $base $sb }" & vbLf
    s = s & "            'Canvas' {" & vbLf
    s = s & "                LoadResources $c $base" & vbLf
    s = s & "                $m = $null" & vbLf
    s = s & "                $attrs = CommonAttrs $c 'Canvas' $ctm ([ref]$m) $base" & vbLf
    s = s & "                $inner = New-Object System.Text.StringBuilder" & vbLf
    s = s & "                DoChildren $c $m $base $inner" & vbLf
    s = s & "                if ($inner.Length -gt 0) { [void]$sb.Append('<g' + $attrs + '>' + $inner.ToString() + '</g>') }" & vbLf
    s = s & "            }" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "# ------------------------------------------------------------------ main" & vbLf
    s = s & "try {" & vbLf
    s = s & "    Report 8 'open'" & vbLf
    s = s & "    $page = FirstPage" & vbLf
    s = s & "    $doc = ReadXml $page" & vbLf
    s = s & "    if (-not $doc) { throw ""cannot read $page"" }" & vbLf
    s = s & "    $root = $doc.DocumentElement" & vbLf
    s = s & "    $script:PageW = Num $root.GetAttribute('Width')" & vbLf
    s = s & "    $script:PageH = Num $root.GetAttribute('Height')" & vbLf
    s = s & "    $script:Total = [Math]::Max(1, $doc.SelectNodes(""//*[local-name()='Path' or local-name()='Glyphs']"").Count)" & vbLf
    s = s & "    Report 12 'page' 0 $script:Total" & vbLf
    s = s & "    LoadResources $root $page" & vbLf
    s = s & "    DoChildren $root @(1.0, 0.0, 0.0, 1.0, 0.0, 0.0) $page $Body" & vbLf
    s = s & "    if ($Body.Length -eq 0) { throw 'empty page' }" & vbLf
    s = s & "    Report 92 'write'" & vbLf
    s = s & "" & vbLf
    s = s & "    # crop: the drawing's frame (from PowerPoint) united with what was actually drawn" & vbLf
    s = s & "    $crop = $null" & vbLf
    s = s & "    if ($Frame -and $Slide) {" & vbLf
    s = s & "        $f = Numbers $Frame; $s = Numbers $Slide" & vbLf
    s = s & "        if ($f.Count -eq 4 -and $s.Count -eq 2 -and $s[0] -gt 0) {" & vbLf
    s = s & "            $k = $script:PageW / $s[0]" & vbLf
    s = s & "            $crop = @(($f[0] * $k), ($f[1] * $k), (($f[0] + $f[2]) * $k), (($f[1] + $f[3]) * $k))" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    if ($script:Box) {" & vbLf
    s = s & "        $b = @([Math]::Max($script:Box[0], 0.0), [Math]::Max($script:Box[1], 0.0)," & vbLf
    s = s & "            [Math]::Min($script:Box[2], $script:PageW), [Math]::Min($script:Box[3], $script:PageH))" & vbLf
    s = s & "        if ($null -eq $crop) { $crop = $b }" & vbLf
    s = s & "        elseif ($b[2] -gt $b[0] -and $b[3] -gt $b[1]) {" & vbLf
    s = s & "            $crop = @([Math]::Min($crop[0], $b[0]), [Math]::Min($crop[1], $b[1]), [Math]::Max($crop[2], $b[2]), [Math]::Max($crop[3], $b[3]))" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    if ($null -eq $crop) { $crop = @(0.0, 0.0, $script:PageW, $script:PageH) }" & vbLf
    s = s & "    $cw = [Math]::Max($crop[2] - $crop[0], 0.1); $ch = [Math]::Max($crop[3] - $crop[1], 0.1)" & vbLf
    s = s & "" & vbLf
    s = s & "    $svg = New-Object System.Text.StringBuilder" & vbLf
    s = s & "    [void]$svg.Append('<?xml version=""1.0"" encoding=""UTF-8""?>' + ""`n"")" & vbLf
    s = s & "    [void]$svg.Append('<svg xmlns=""http://www.w3.org/2000/svg"" xmlns:xlink=""http://www.w3.org/1999/xlink"" version=""1.1"" width=""' + (N ($cw * 0.75)) + 'pt"" height=""' + (N ($ch * 0.75)) + 'pt"" viewBox=""' + (N $crop[0]) + ' ' + (N $crop[1]) + ' ' + (N $cw) + ' ' + (N $ch) + '"">' + ""`n"")" & vbLf
    s = s & "    [void]$svg.Append('<!-- Created by VecStamp (https://github.com/vluckyzhang/VecStamp) from PowerPoint XPS output -->' + ""`n"")" & vbLf
    s = s & "    if ($Defs.Length -gt 0) { [void]$svg.Append('<defs>' + $Defs.ToString() + '</defs>' + ""`n"") }" & vbLf
    s = s & "    [void]$svg.Append($Body.ToString() + ""`n</svg>`n"")" & vbLf
    s = s & "    $enc = New-Object System.Text.UTF8Encoding($false)" & vbLf
    s = s & "    [IO.File]::WriteAllText($Out, $svg.ToString(), $enc)" & vbLf
    s = s & "    $Zip.Dispose()" & vbLf
    s = s & "    $script:LastPct = -1" & vbLf
    s = s & "    Report 100 'done'" & vbLf
    s = s & "    exit 0" & vbLf
    s = s & "}" & vbLf
    s = s & "catch {" & vbLf
    s = s & "    try { $Zip.Dispose() } catch { }" & vbLf
    s = s & "    [Console]::Error.WriteLine('xps2svg: ' + $_.Exception.Message)" & vbLf
    s = s & "    if ($_.Exception.Message -eq 'timeout') { exit 3 }" & vbLf
    s = s & "    exit 1" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    HelperScript5 = s
End Function

Private Function HelperScript() As String
    HelperScript = HelperScript1() & HelperScript2() & HelperScript3() & HelperScript4() & HelperScript5()
End Function

Private Function PresetSrc1(ByVal ast As Long) As String
    Dim s As String
    Select Case ast
        Case 1      ' rect
            s = "||f,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z|l,t,r,b"
        Case 2      ' parallelogram
            s = "25000|m,100000,w,ss;p,0,@1,$0;m,ss,$1,200000;m,ss,$1,100000;a,r,0,$2;a,r,0,$3;m,$5,1,2;a,r,0,$6;m,wd2,$1,$0;m,5,$1,$0;d,1,$9,12;m,$10,w,1;m,$10,h,1;a,r,0,$11;a,b,0,$12;m,h,hc,$3;p,0,$15,h;a,b,0,$16|f,1,0,0:M,l,b;L,$3,t;L,r,t;L,$5,b;Z|$11,$12,$13,$14"
        Case 3      ' trapezoid
            s = "25000|m,50000,w,ss;p,0,@1,$0;m,ss,$1,200000;m,ss,$1,100000;a,r,0,$3;a,r,0,$2;m,wd3,$1,$0;m,hd3,$1,$0;a,r,0,$6|f,1,0,0:M,l,b;L,$3,t;L,$4,t;L,r,b;Z|$6,$7,$8,b"
        Case 4      ' diamond
            s = "|m,w,3,4;m,h,3,4|f,1,0,0:M,l,vc;L,hc,t;L,r,vc;L,hc,b;Z|wd4,hd4,$0,$1"
        Case 5      ' roundRect
            s = "16667|p,0,@1,50000;m,ss,$0,100000;a,r,0,$1;a,b,0,$1;m,$1,29289,100000;a,r,0,$4;a,b,0,$4|f,1,0,0:M,l,$1;A,$1,$1,cd2,cd4;L,$2,t;A,$1,$1,k34,cd4;L,r,$3;A,$1,$1,0,cd4;L,$1,b;A,$1,$1,cd4,cd4;Z|$4,$4,$5,$6"
        Case 6      ' octagon
            s = "29289|p,0,@1,50000;m,ss,$0,100000;a,r,0,$1;a,b,0,$1;m,$1,1,2;a,r,0,$4;a,b,0,$4|f,1,0,0:M,l,$1;L,$1,t;L,$2,t;L,r,$1;L,r,$3;L,$2,b;L,$1,b;L,l,$3;Z|$4,$4,$5,$6"
        Case 7      ' triangle
            s = "50000|p,0,@1,100000;m,w,$0,200000;m,w,$0,100000;a,$1,wd2,0|f,1,0,0:M,l,b;L,$2,t;L,r,b;Z|$1,vc,$3,b"
        Case 8      ' rtTriangle
            s = "|m,h,7,12;m,w,7,12;m,h,11,12|f,1,0,0:M,l,b;L,l,t;L,r,b;Z|wd12,$0,$1,$2"
        Case 9      ' ellipse
            s = "|o,wd2,2700000;i,hd2,2700000;a,hc,0,$0;a,hc,$0,0;a,vc,0,$1;a,vc,$1,0|f,1,0,0:M,l,vc;A,wd2,hd2,cd2,cd4;A,wd2,hd2,k34,cd4;A,wd2,hd2,0,cd4;A,wd2,hd2,cd4,cd4;Z|$2,$4,$3,$5"
        Case 10      ' hexagon
            s = "25000,115470|m,50000,w,ss;p,0,@1,$0;m,hd2,@2,100000;m,ss,$1,100000;a,r,0,$3;i,$2,3600000;a,vc,0,$5;a,vc,$5,0;m,$0,-1,2;a,$1,$8,0;q,$9,4,2;q,$9,3,2;q,$9,$8,0;d,$1,$12,$8;m,$13,$11,-1;a,$10,$14,0;m,w,$15,24;m,h,$15,24;a,r,0,$16;a,b,0,$17|f,1,0,0:M,l,vc;L,$3,$6;L,$4,$6;L,r,vc;L,$4,$7;L,$3,$7;Z|$16,$17,$18,$19"
        Case 11      ' plus
            s = "25000|p,0,@1,50000;m,ss,$0,100000;a,r,0,$1;a,b,0,$1;a,w,0,h;q,$4,l,$1;q,$4,r,$2;q,$4,$1,t;q,$4,$3,b|f,1,0,0:M,l,$1;L,$1,$1;L,$1,t;L,$2,t;L,$2,$1;L,r,$1;L,r,$3;L,$2,$3;L,$2,b;L,$1,b;L,$1,$3;L,l,$3;Z|$5,$7,$6,$8"
        Case 12      ' pentagon
            s = "105146,110557|m,wd2,@1,100000;m,hd2,@2,100000;m,vc,@2,100000;o,$0,1080000;o,$0,18360000;i,$1,1080000;i,$1,18360000;a,hc,0,$3;a,hc,0,$4;a,hc,$4,0;a,hc,$3,0;a,$2,0,$5;a,$2,0,$6;m,$11,$4,$3|f,1,0,0:M,$7,$11;L,hc,t;L,$10,$11;L,$9,$12;L,$8,$12;Z|$8,$13,$9,$12"
        Case 13      ' can
            s = "25000|m,50000,h,ss;p,0,@1,$0;m,ss,$1,200000;a,$2,$2,0;a,b,0,$2|f,0,0,0:M,l,$2;A,wd2,$2,cd2,-10800000;L,r,$4;A,wd2,$2,0,cd2;Z~f,0,0,0:M,l,$2;A,wd2,$2,cd2,cd2;A,wd2,$2,0,cd2;Z~n,1,0,0:M,r,$2;A,wd2,$2,0,cd2;A,wd2,$2,cd2,cd2;L,r,$4;A,wd2,$2,0,cd2;L,l,$2|l,$3,r,$4"
        Case 14      ' cube
            s = "25000|p,0,@1,100000;m,ss,$0,100000;a,b,0,$1;m,$2,1,2;d,$1,b,2;a,r,0,$1;m,$5,1,2;d,$1,r,2|f,0,0,0:M,l,$1;L,$5,$1;L,$5,b;L,l,b;Z~f,0,0,0:M,$5,$1;L,r,t;L,r,$2;L,$5,b;Z~f,0,0,0:M,l,$1;L,$1,t;L,r,t;L,$5,$1;Z~n,1,0,0:M,l,$1;L,$1,t;L,r,t;L,r,$2;L,$5,b;L,l,b;Z;M,l,$1;L,$5,$1;L,r,t;M,$5,$1;L,$5,b|l,$1,$5,b"
        Case 15      ' bevel
            s = "12500|p,0,@1,50000;m,ss,$0,100000;a,r,0,$1;a,b,0,$1|f,0,0,0:M,$1,$1;L,$2,$1;L,$2,$3;L,$1,$3;Z~f,0,0,0:M,l,t;L,r,t;L,$2,$1;L,$1,$1;Z~f,0,0,0:M,l,b;L,$1,$3;L,$2,$3;L,r,b;Z~f,0,0,0:M,l,t;L,$1,$1;L,$1,$3;L,l,b;Z~f,0,0,0:M,r,t;L,r,b;L,$2,$3;L,$2,$1;Z~n,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z;M,$1,$1;L,$2,$1;L,$2,$3;L,$1,$3;Z;M,l,t;L,$1,$1;M,l,b;L,$1,$3;M,r,t;L,$2,$1;M,r,b;L,$2,$3|$1,$1,$2,$3"
        Case 16      ' foldedCorner
            s = "16667|p,0,@1,50000;m,ss,$0,100000;m,$1,1,5;a,r,0,$1;a,$3,$2,0;a,b,0,$1;a,$5,$2,0|f,0,0,0:M,l,t;L,r,t;L,r,$5;L,$3,b;L,l,b;Z~f,0,0,0:M,$3,b;L,$4,$6;L,r,$5;Z~n,1,0,0:M,$3,b;L,$4,$6;L,r,$5;L,$3,b;L,l,b;L,l,t;L,r,t;L,r,$5|l,t,r,$5"
        Case 17      ' smileyFace
            s = "4653|p,-4653,@1,4653;m,w,4969,21699;m,w,6215,21600;m,w,13135,21600;m,w,16640,21600;m,h,7570,21600;m,h,16515,21600;m,h,$0,100000;a,$6,0,$7;a,$6,$7,0;m,h,$0,50000;a,$9,$10,0;o,wd2,2700000;i,hd2,2700000;a,hc,0,$12;a,hc,$12,0;a,vc,0,$13;a,vc,$13,0;m,w,1125,21600;m,h,1125,21600|f,0,0,0:M,l,vc;A,wd2,hd2,cd2,21600000;Z~f,1,0,0:M,$2,$5;A,$18,$19,cd2,21600000;M,$3,$5;A,$18,$19,cd2,21600000~n,1,0,0:M,$1,$8;Q,hc,$11,$4,$8~n,1,0,0:M,l,vc;A,wd2,hd2,cd2,21600000;Z|$14,$16,$15,$17"
        Case 18      ' donut
            s = "25000|p,0,@1,50000;m,ss,$0,100000;a,wd2,0,$1;a,hd2,0,$1;o,wd2,2700000;i,hd2,2700000;a,hc,0,$4;a,hc,$4,0;a,vc,0,$5;a,vc,$5,0|f,1,0,0:M,l,vc;A,wd2,hd2,cd2,cd4;A,wd2,hd2,k34,cd4;A,wd2,hd2,0,cd4;A,wd2,hd2,cd4,cd4;Z;M,$1,vc;A,$2,$3,cd2,-5400000;A,$2,$3,cd4,-5400000;A,$2,$3,0,-5400000;A,$2,$3,k34,-5400000;Z|$6,$8,$7,$9"
        Case 19      ' noSmoking
            s = "18750|p,0,@1,50000;m,ss,$0,100000;a,wd2,0,$1;a,hd2,0,$1;t,w,h;o,$3,$4;i,$2,$4;h,$5,$6,0;m,$2,$3,$7;m,$1,1,2;t,$8,$9;m,$10,2,1;a,-10800000,$11,0;t,w,h;a,$13,0,$10;a,$14,0,cd2;o,$3,$14;i,$2,$14;h,$16,$17,0;m,$2,$3,$18;o,$19,$14;i,$19,$14;a,hc,$20,0;a,vc,$21,0;a,hc,0,$20;a,vc,0,$21;o,wd2,2700000;i,hd2,2700000;a,hc,0,$26;a,hc,$26,0;a,vc,0,$27;a,vc,$27,0|f,1,0,0:M,l,vc;A,wd2,hd2,cd2,cd4;A,wd2,hd2,k34,cd4;A,wd2,hd2,0,cd4;A,wd2,hd2,cd4,cd4;Z;M,$22,$23;A,$2,$3,$14,$12;Z;M,$24,$25;A,$2,$3,$15,$12;Z|$28,$30,$29,$31"
        Case 20      ' blockArc
            s = "!10800000,!0,25000|p,0,@1,21599999;p,0,@2,21599999;p,0,@3,50000;a,$1,0,$0;a,$3,21600000,0;q,$3,$3,$4;a,0,0,$5;i,wd2,$0;o,hd2,$0;i,wd2,$1;o,hd2,$1;c,wd2,$8,$7;s,hd2,$8,$7;c,wd2,$10,$9;s,hd2,$10,$9;a,hc,$11,0;a,vc,$12,0;a,hc,$13,0;a,vc,$14,0;m,ss,$2,100000;a,wd2,0,$19;a,hd2,0,$19;i,$20,$1;o,$21,$1;i,$20,$0;o,$21,$0;c,$20,$23,$22;s,$21,$23,$22;c,$20,$25,$24;s,$21,$25,$24;a,hc,$26,0;a,vc,$27,0;a,hc,$28,0;a,vc,$29,0;a,21600000,0,$0;a,$5,0,$34;x,$15,$30;x,$17,$32;x,$36,$37;q,$35,r,$38;a,cd4,0,$0;a,27000000,0,$0;q,$40,$40,$41;a,$5,0,$42;x,$16,$31;x,$18,$33;x,$44,$45;q,$43,b,$46;a,cd2,0,$0;a,32400000,0,$0;q,$48,$48,$49;a,$5,0,$50;n,$15,$30;n,$17,$32;n,$52,$53;q,$51,l,$54;a,k34,0,$0;a,37800000,0,$0;q,$56,$56,$57;a,$5,0,$58;n,$16,$31;n,$18,$33;n,$60,$61;q,$59,t,$62;d,$15,$32,2;d,$16,$33,2;d,$17,$30,2;d,$18,$31,2;a,$0,0,cd4;a,$1,cd4,0;d,$68,$69,2|f,1,0,0:M,$15,$16;A,wd2,hd2,$0,$5;L,$30,$31;A,$20,$2"
            s = s & "1,$1,$6;Z|$55,$63,$39,$47"
        Case 21      ' heart
            s = "|m,w,49,48;m,w,10,48;a,hc,0,$0;a,hc,0,$1;a,hc,$1,0;a,hc,$0,0;a,t,0,hd3;m,w,1,6;m,w,5,6;m,h,2,3|f,1,0,0:M,hc,hd4;C,$4,$6,$5,hd4,hc,b;C,$2,hd4,$3,$6,hc,hd4;Z|$7,hd4,$8,$9"
        Case 22      ' lightningBolt
            s = "|m,w,5022,21600;m,w,8472,21600;m,w,8757,21600;m,w,10012,21600;m,w,12860,21600;m,w,13917,21600;m,w,16577,21600;m,h,3890,21600;m,h,6080,21600;m,h,7437,21600;m,h,9705,21600;m,h,12007,21600;m,h,14277,21600;m,h,14915,21600|f,1,21600,21600:M,8472,0;L,12860,6080;L,11050,6797;L,16577,12007;L,14767,12877;L,21600,21600;L,10012,14915;L,12222,13987;L,5022,9705;L,7602,8382;L,0,3890;Z|$2,$9,$5,$12"
        Case 23      ' sun
            s = "25000|p,12500,@1,46875;a,50000,0,$0;m,$1,30274,32768;m,$1,12540,32768;a,$2,50000,0;a,$3,50000,0;a,50000,0,$2;a,50000,0,$3;m,$1,23170,32768;a,50000,$8,0;a,50000,0,$8;m,$6,3,4;m,$7,3,4;a,$11,3662,0;a,$12,3662,0;a,$12,12500,0;a,100000,0,$11;a,100000,0,$13;a,100000,0,$14;a,100000,0,$15;m,w,18436,21600;m,h,3163,21600;m,w,3163,21600;m,h,18436,21600;m,w,$9,100000;m,w,$10,100000;m,w,$11,100000;m,w,$13,100000;m,w,$14,100000;m,w,$15,100000;m,w,$16,100000;m,w,$17,100000;m,w,$18,100000;m,w,$19,100000;m,w,$0,100000;m,w,$1,100000;m,h,$1,100000;m,h,$9,100000;m,h,$10,100000;m,h,$11,100000;m,h,$13,100000;m,h,$14,100000;m,h,$15,100000;m,h,$16,100000;m,h,$17,100000;m,h,$18,100000;m,h,$19,100000|f,1,0,0:M,r,vc;L,$30,$46;L,$30,$42;Z;M,$20,$21;L,$31,$41;L,$32,$40;Z;M,hc,t;L,$33,$39;L,$29,$39;Z;M,$22,$21;L,$28,$40;L,$27,$41;Z;M,l,vc;L,$26,$42;L,$26,$46;Z;M,$22,$23;L,$27,$45;L,$28,$44;Z;M,hc,b;L,$29,$43;L,$33,$"
            s = s & "43;Z;M,$20,$23;L,$32,$44;L,$31,$45;Z;M,$34,vc;A,$35,$36,cd2,21600000;Z|$25,$38,$24,$37"
        Case 24      ' moon
            s = "50000|p,0,@1,87500;m,ss,$0,100000;m,$1,w,ss;a,ss,0,$1;m,$1,$1,$3;m,ss,ss,$3;m,$5,2,1;a,$6,0,$4;a,$7,0,$1;m,$8,w,ss;m,$7,1,2;a,$10,0,$1;m,$11,hd2,ss;a,vc,0,$12;a,vc,$12,0;m,$1,9598,32768;m,$15,w,ss;a,ss,0,$15;m,ss,ss,1;m,$17,$17,1;a,$18,0,$19;r,$20;m,$21,hd2,ss;a,vc,0,$22;a,vc,$22,0;a,$9,0,$2;m,$25,1,2;a,$2,$26,w;m,$27,-1,1;m,hd2,-1,1;t,$28,$29;t,$28,hd2;a,$31,0,21600000;a,$32,0,$30|f,1,0,0:M,r,b;A,w,hd2,cd4,cd2;A,$26,$12,$30,$33;Z|$16,$23,$2,$24"
        Case 25      ' arc
            s = "!16200000,!0|p,0,@1,21599999;p,0,@2,21599999;a,$1,0,$0;a,$2,21600000,0;q,$2,$2,$3;i,wd2,$0;o,hd2,$0;c,wd2,$6,$5;s,hd2,$6,$5;i,wd2,$1;o,hd2,$1;c,wd2,$10,$9;s,hd2,$10,$9;a,hc,$7,0;a,vc,$8,0;a,hc,$11,0;a,vc,$12,0;a,21600000,0,$0;a,$4,0,$17;x,$13,$15;q,$18,r,$19;a,cd4,0,$0;a,27000000,0,$0;q,$21,$21,$22;a,$4,0,$23;x,$14,$16;q,$24,b,$25;a,cd2,0,$0;a,32400000,0,$0;q,$27,$27,$28;a,$4,0,$29;n,$13,$15;q,$30,l,$31;a,k34,0,$0;a,37800000,0,$0;q,$33,$33,$34;a,$4,0,$35;n,$14,$16;q,$36,t,$37;a,$0,0,cd4;a,$1,cd4,0;d,$39,$40,2|f,0,0,0:M,$13,$14;A,wd2,hd2,$0,$4;L,hc,vc;Z~n,1,0,0:M,$13,$14;A,wd2,hd2,$0,$4|$32,$38,$20,$26"
        Case 26      ' bracketPair
            s = "16667|p,0,@1,50000;m,ss,$0,100000;a,r,0,$1;a,b,0,$1;m,$1,29289,100000;a,r,0,$4;a,b,0,$4|f,0,0,0:M,l,$1;A,$1,$1,cd2,cd4;L,$2,t;A,$1,$1,k34,cd4;L,r,$3;A,$1,$1,0,cd4;L,$1,b;A,$1,$1,cd4,cd4;Z~n,1,0,0:M,$1,b;A,$1,$1,cd4,cd4;L,l,$1;A,$1,$1,cd2,cd4;M,$2,t;A,$1,$1,k34,cd4;L,r,$3;A,$1,$1,0,cd4|$4,$4,$5,$6"
        Case 27      ' bracePair
            s = "8333|p,0,@1,25000;m,ss,$0,100000;m,ss,$0,50000;a,r,0,$2;a,r,0,$1;a,vc,0,$1;a,vc,$1,0;a,b,0,$1;m,$1,29289,100000;a,$1,$8,0;a,r,0,$9;a,b,0,$8|f,0,0,0:M,$2,b;A,$1,$1,cd4,cd4;L,$1,$6;A,$1,$1,0,-5400000;A,$1,$1,cd4,-5400000;L,$1,$1;A,$1,$1,cd2,cd4;L,$3,t;A,$1,$1,k34,cd4;L,$4,$5;A,$1,$1,cd2,-5400000;A,$1,$1,k34,-5400000;L,$4,$7;A,$1,$1,0,cd4;Z~n,1,0,0:M,$2,b;A,$1,$1,cd4,cd4;L,$1,$6;A,$1,$1,0,-5400000;A,$1,$1,cd4,-5400000;L,$1,$1;A,$1,$1,cd2,cd4;M,$3,t;A,$1,$1,k34,cd4;L,$4,$5;A,$1,$1,cd2,-5400000;A,$1,$1,k34,-5400000;L,$4,$7;A,$1,$1,0,cd4|$9,$9,$10,$11"
        Case 28      ' plaque
            s = "16667|p,0,@1,50000;m,ss,$0,100000;a,r,0,$1;a,b,0,$1;m,$1,70711,100000;a,r,0,$4;a,b,0,$4|f,1,0,0:M,l,$1;A,$1,$1,cd4,-5400000;L,$2,t;A,$1,$1,cd2,-5400000;L,r,$3;A,$1,$1,k34,-5400000;L,$1,b;A,$1,$1,0,-5400000;Z|$4,$4,$5,$6"
        Case 29      ' leftBracket
            s = "8333|m,50000,h,ss;p,0,@1,$0;m,ss,$1,100000;a,b,0,$2;o,w,2700000;i,$2,2700000;a,r,0,$4;a,$2,0,$5;a,b,$5,$2|f,0,0,0:M,r,b;A,w,$2,cd4,cd4;L,l,$2;A,w,$2,cd2,cd4;Z~n,1,0,0:M,r,b;A,w,$2,cd4,cd4;L,l,$2;A,w,$2,cd2,cd4|$6,$7,r,$8"
        Case 30      ' rightBracket
            s = "8333|m,50000,h,ss;p,0,@1,$0;m,ss,$1,100000;a,b,0,$2;o,w,2700000;i,$2,2700000;a,l,$4,0;a,$2,0,$5;a,b,$5,$2|f,0,0,0:M,l,t;A,w,$2,k34,cd4;L,r,$3;A,w,$2,0,cd4;Z~n,1,0,0:M,l,t;A,w,$2,k34,cd4;L,r,$3;A,w,$2,0,cd4|l,$7,$6,$8"
        Case 31      ' leftBrace
            s = "8333,50000|p,0,@2,100000;a,100000,0,$0;n,$1,$0;m,$2,1,2;m,$3,h,ss;p,0,@1,$4;m,ss,$5,100000;m,h,$0,100000;a,$7,$6,0;o,wd2,2700000;i,$6,2700000;a,r,0,$9;a,$6,0,$10;a,b,$10,$6|f,0,0,0:M,r,b;A,wd2,$6,cd4,cd4;L,hc,$8;A,wd2,$6,0,-5400000;A,wd2,$6,cd4,-5400000;L,hc,$6;A,wd2,$6,cd2,cd4;Z~n,1,0,0:M,r,b;A,wd2,$6,cd4,cd4;L,hc,$8;A,wd2,$6,0,-5400000;A,wd2,$6,cd4,-5400000;L,hc,$6;A,wd2,$6,cd2,cd4|$11,$12,r,$13"
        Case 32      ' rightBrace
            s = "8333,50000|p,0,@2,100000;a,100000,0,$0;n,$1,$0;m,$2,1,2;m,$3,h,ss;p,0,@1,$4;m,ss,$5,100000;m,h,$0,100000;a,$7,0,$6;a,b,0,$6;o,wd2,2700000;i,$6,2700000;a,l,$10,0;a,$6,0,$11;a,b,$11,$6|f,0,0,0:M,l,t;A,wd2,$6,k34,cd4;L,hc,$8;A,wd2,$6,cd2,-5400000;A,wd2,$6,k34,-5400000;L,hc,$9;A,wd2,$6,0,cd4;Z~n,1,0,0:M,l,t;A,wd2,$6,k34,cd4;L,hc,$8;A,wd2,$6,cd2,-5400000;A,wd2,$6,k34,-5400000;L,hc,$9;A,wd2,$6,0,cd4|l,$13,$12,$14"
        Case 33      ' rightArrow
            s = "50000,50000|m,100000,w,ss;p,0,@1,100000;p,0,@2,$0;m,ss,$2,100000;a,r,0,$3;m,h,$1,200000;a,vc,0,$5;a,vc,$5,0;m,$6,$3,hd2;a,$4,$8,0|f,1,0,0:M,l,$6;L,$4,$6;L,$4,t;L,r,vc;L,$4,b;L,$4,$7;L,l,$7;Z|l,$6,$9,$7"
        Case 34      ' leftArrow
            s = "50000,50000|m,100000,w,ss;p,0,@1,100000;p,0,@2,$0;m,ss,$2,100000;a,l,$3,0;m,h,$1,200000;a,vc,0,$5;a,vc,$5,0;m,$6,$3,hd2;a,$4,0,$8|f,1,0,0:M,l,vc;L,$4,t;L,$4,$6;L,r,$6;L,r,$7;L,$4,$7;L,$4,b;Z|$9,$6,r,$7"
        Case 35      ' upArrow
            s = "50000,50000|m,100000,h,ss;p,0,@1,100000;p,0,@2,$0;m,ss,$2,100000;a,t,$3,0;m,w,$1,200000;a,hc,0,$5;a,hc,$5,0;m,$6,$3,wd2;a,$4,0,$8|f,1,0,0:M,l,$4;L,hc,t;L,r,$4;L,$7,$4;L,$7,b;L,$6,b;L,$6,$4;Z|$6,$9,$7,b"
        Case 36      ' downArrow
            s = "50000,50000|m,100000,h,ss;p,0,@1,100000;p,0,@2,$0;m,ss,$2,100000;a,b,0,$3;m,w,$1,200000;a,hc,0,$5;a,hc,$5,0;m,$6,$3,wd2;a,$4,$8,0|f,1,0,0:M,l,$4;L,$6,$4;L,$6,t;L,$7,t;L,$7,$4;L,r,$4;L,hc,b;Z|$6,t,$7,$9"
        Case 37      ' leftRightArrow
            s = "50000,50000|m,50000,w,ss;p,0,@1,100000;p,0,@2,$0;m,ss,$2,100000;a,r,0,$3;m,h,$1,200000;a,vc,0,$5;a,vc,$5,0;m,$6,$3,hd2;a,$3,0,$8;a,$4,$8,0|f,1,0,0:M,l,vc;L,$3,t;L,$3,$6;L,$4,$6;L,$4,t;L,r,vc;L,$4,b;L,$4,$7;L,$3,$7;L,$3,b;Z|$9,$6,$10,$7"
        Case 38      ' upDownArrow
            s = "50000,50000|m,50000,h,ss;p,0,@1,100000;p,0,@2,$0;m,ss,$2,100000;a,b,0,$3;m,w,$1,200000;a,hc,0,$5;a,hc,$5,0;m,$6,$3,wd2;a,$3,0,$8;a,$4,$8,0|f,1,0,0:M,l,$3;L,hc,t;L,r,$3;L,$7,$3;L,$7,$4;L,r,$4;L,hc,b;L,l,$4;L,$6,$4;L,$6,$3;Z|$6,$9,$7,$10"
        Case 39      ' quadArrow
            s = "22500,22500,22500|p,0,@2,50000;m,$0,2,1;p,0,@1,$1;a,100000,0,$1;m,$3,1,2;p,0,@3,$4;m,ss,$5,100000;m,ss,$0,100000;a,hc,0,$7;a,hc,$7,0;m,ss,$2,200000;a,hc,0,$10;a,hc,$10,0;a,r,0,$6;a,vc,0,$7;a,vc,$7,0;a,vc,0,$10;a,vc,$10,0;a,b,0,$6;m,$10,$6,$7;a,r,0,$19|f,1,0,0:M,l,vc;L,$6,$14;L,$6,$16;L,$11,$16;L,$11,$6;L,$8,$6;L,hc,t;L,$9,$6;L,$12,$6;L,$12,$16;L,$13,$16;L,$13,$14;L,r,vc;L,$13,$15;L,$13,$17;L,$12,$17;L,$12,$18;L,$9,$18;L,hc,b;L,$8,$18;L,$11,$18;L,$11,$17;L,$6,$17;L,$6,$15;Z|$19,$16,$20,$17"
        Case 40      ' leftRightUpArrow
            s = "25000,25000,25000|p,0,@2,50000;m,$0,2,1;p,0,@1,$1;a,100000,0,$1;m,$3,1,2;p,0,@3,$4;m,ss,$5,100000;m,ss,$0,100000;a,hc,0,$7;a,hc,$7,0;m,ss,$2,200000;a,hc,0,$10;a,hc,$10,0;a,r,0,$6;m,ss,$0,50000;a,b,0,$14;a,b,0,$7;a,$16,0,$10;a,$16,$10,0;m,$10,$6,$7;a,r,0,$19|f,1,0,0:M,l,$16;L,$6,$15;L,$6,$17;L,$11,$17;L,$11,$6;L,$8,$6;L,hc,t;L,$9,$6;L,$12,$6;L,$12,$17;L,$13,$17;L,$13,$15;L,r,$16;L,$13,b;L,$13,$18;L,$6,$18;L,$6,b;Z|$19,$17,$20,$18"
        Case 41      ' bentArrow
            s = "25000,25000,25000,43750|p,0,@2,50000;m,$0,2,1;p,0,@1,$1;p,0,@3,50000;m,ss,$2,100000;m,ss,$0,100000;m,$4,1,2;a,$5,0,$6;m,ss,$3,100000;a,r,0,$8;a,b,0,$7;n,$9,$10;m,100000,$11,ss;p,0,@4,$12;m,ss,$13,100000;a,$14,0,$4;x,$15,0;a,$4,$16,0;a,r,0,$8;a,$7,$4,0;a,$19,$7,0;a,$7,$14,0;a,$19,$16,0|f,1,0,0:M,l,b;L,l,$21;A,$14,$14,cd2,cd4;L,$18,$7;L,$18,t;L,r,$5;L,$18,$20;L,$18,$19;L,$17,$19;A,$16,$16,k34,-5400000;L,$4,b;Z|l,t,r,b"
        Case 42      ' uturnArrow
            s = "25000,25000,25000,43750,75000|p,0,@2,25000;m,$0,2,1;p,0,@1,$1;m,$2,ss,h;a,100000,0,$3;m,$4,h,ss;p,0,@3,$5;a,$6,$2,0;m,$7,ss,h;p,$8,@5,100000;m,ss,$2,100000;m,ss,$0,100000;m,$10,1,2;a,$11,0,$12;m,h,$9,100000;m,ss,$6,100000;a,$14,0,$15;a,r,0,$13;m,$17,1,2;n,$18,$16;m,$19,100000,ss;p,0,@4,$20;m,ss,$21,100000;a,$22,0,$10;x,$23,0;a,$10,$24,0;a,r,0,$11;a,$26,0,$11;a,$27,$13,0;a,$17,0,$22;a,$28,0,$24;d,$10,$28,2|f,1,0,0:M,l,b;L,l,$22;A,$22,$22,cd2,cd4;L,$29,t;A,$22,$22,k34,cd4;L,$17,$16;L,r,$16;L,$26,$14;L,$27,$16;L,$28,$16;L,$28,$25;A,$24,$24,0,-5400000;L,$25,$10;A,$24,$24,k34,-5400000;L,$10,b;Z|l,t,r,b"
        Case 43      ' leftUpArrow
            s = "25000,25000,25000|p,0,@2,50000;m,$0,2,1;p,0,@1,$1;a,100000,0,$1;p,0,@3,$3;m,ss,$4,100000;m,ss,$0,50000;a,r,0,$6;a,b,0,$6;m,ss,$0,100000;a,r,0,$9;a,b,0,$9;m,ss,$2,200000;a,$10,0,$12;a,$10,$12,0;a,$11,0,$12;a,$11,$12,0;m,$12,$5,$9;d,$5,$14,2;d,$5,$16,2|f,1,0,0:M,l,$11;L,$5,$8;L,$5,$15;L,$13,$15;L,$13,$5;L,$7,$5;L,$10,t;L,r,$5;L,$14,$5;L,$14,$16;L,$5,$16;L,$5,b;Z|$17,$15,$10,$16"
        Case 44      ' bentUpArrow
            s = "25000,25000,25000|p,0,@1,50000;p,0,@2,50000;p,0,@3,50000;m,ss,$2,100000;m,ss,$1,50000;a,r,0,$4;m,ss,$1,100000;a,r,0,$6;m,ss,$0,200000;a,$7,0,$8;a,$7,$8,0;m,ss,$0,100000;a,b,0,$11;m,$10,1,2;d,$12,b,2;d,$3,b,2|f,1,0,0:M,l,$12;L,$9,$12;L,$9,$3;L,$5,$3;L,$7,t;L,r,$3;L,$10,$3;L,$10,b;L,l,b;Z|l,$12,$10,b"
        Case 45      ' curvedRightArrow
            s = "25000,50000,25000|m,50000,h,ss;p,0,@2,$0;p,0,@1,$1;m,ss,$2,100000;m,ss,$1,100000;d,$3,$4,4;a,hd2,0,$5;m,$6,2,1;m,$7,$7,1;m,$3,$3,1;a,$8,0,$9;r,$10;m,$11,w,$7;m,100000,$12,ss;p,0,@3,$13;m,ss,$14,100000;a,$6,$3,0;m,w,w,1;m,$15,$15,1;a,$17,0,$18;r,$19;m,$20,$6,w;a,$6,$21,0;a,$16,$21,0;a,$4,0,$3;m,$24,1,2;a,$22,0,$25;a,$23,$25,0;m,$4,1,2;a,b,0,$28;a,r,0,$15;t,$15,$21;a,cd2,0,$31;a,0,0,$31;a,r,0,$12;d,$6,$16,2;m,$3,1,2;t,$12,$36;a,$37,0,cd4;a,cd4,$37,0;a,cd2,0,$37|f,0,0,0:M,l,$6;A,w,$6,cd2,$33;L,$30,$26;L,r,$29;L,$30,$27;L,$30,$23;A,w,$6,$32,$31;Z~f,0,0,0:M,r,$3;A,w,$6,k34,$38;A,w,$6,$40,$39;Z~n,1,0,0:M,l,$6;A,w,$6,cd2,$33;L,$30,$26;L,r,$29;L,$30,$27;L,$30,$23;A,w,$6,$32,$31;L,l,$6;A,w,$6,cd2,cd4;L,r,$3;A,w,$6,k34,$38|l,t,r,b"
        Case 46      ' curvedLeftArrow
            s = "25000,50000,25000|m,50000,h,ss;p,0,@2,$0;p,0,@1,$1;m,ss,$2,100000;m,ss,$1,100000;d,$3,$4,4;a,hd2,0,$5;m,$6,2,1;m,$7,$7,1;m,$3,$3,1;a,$8,0,$9;r,$10;m,$11,w,$7;m,100000,$12,ss;p,0,@3,$13;m,ss,$14,100000;a,$6,$3,0;m,w,w,1;m,$15,$15,1;a,$17,0,$18;r,$19;m,$20,$6,w;a,$6,$21,0;a,$16,$21,0;a,$4,0,$3;m,$24,1,2;a,$22,0,$25;a,$23,$25,0;m,$4,1,2;a,b,0,$28;a,l,$15,0;t,$15,$21;a,0,0,$31;a,l,$12,0;d,$6,$16,2;m,$3,1,2;t,$12,$35;a,$36,0,$31;a,$31,$36,0;a,0,0,$36|f,0,0,0:M,l,$29;L,$30,$26;L,$30,$22;A,w,$6,$31,$37;A,w,$6,$39,$38;L,$30,$27;Z~f,0,0,0:M,r,$16;A,w,$6,0,-5400000;L,l,t;A,w,$6,k34,cd4;Z~n,1,0,0:M,r,$16;A,w,$6,0,-5400000;L,l,t;A,w,$6,k34,cd4;L,r,$16;A,w,$6,0,$31;L,$30,$27;L,l,$29;L,$30,$26;L,$30,$22;A,w,$6,$31,$37|l,t,r,b"
        Case 47      ' curvedUpArrow
            s = "25000,50000,25000|m,50000,w,ss;p,0,@2,$0;p,0,@1,100000;m,ss,$2,100000;m,ss,$1,100000;d,$3,$4,4;a,wd2,0,$5;m,$6,2,1;m,$7,$7,1;m,$3,$3,1;a,$8,0,$9;r,$10;m,$11,h,$7;m,100000,$12,ss;p,0,@3,$13;m,ss,@3,100000;a,$6,$3,0;m,h,h,1;m,$15,$15,1;a,$17,0,$18;r,$19;m,$20,$6,h;a,$6,$21,0;a,$16,$21,0;a,$4,0,$3;m,$24,1,2;a,$22,0,$25;a,$23,$25,0;m,$4,1,2;a,r,0,$28;a,t,$15,0;t,$15,$21;a,0,0,$31;a,t,$12,0;d,$6,$16,2;m,$3,1,2;t,$12,$35;a,$36,0,$31;a,0,0,$37;a,cd4,0,$31;a,$31,$36,0;a,cd4,0,$36|f,0,0,0:M,$29,t;L,$27,$30;L,$23,$30;A,$6,h,$39,$40;A,$6,h,$41,$37;L,$26,$30;Z~f,0,0,0:M,$6,b;A,$6,h,cd4,cd4;L,$3,t;A,$6,h,cd2,-5400000;Z~n,1,0,0:M,$34,$33;A,$6,h,$41,$37;L,$26,$30;L,$29,t;L,$27,$30;L,$23,$30;A,$6,h,$39,$31;L,$6,b;A,$6,h,cd4,cd4;L,$3,t;A,$6,h,cd2,-5400000|l,t,r,b"
        Case 48      ' curvedDownArrow
            s = "25000,50000,25000|m,50000,w,ss;p,0,@2,$0;p,0,@1,100000;m,ss,$2,100000;m,ss,$1,100000;d,$3,$4,4;a,wd2,0,$5;m,$6,2,1;m,$7,$7,1;m,$3,$3,1;a,$8,0,$9;r,$10;m,$11,h,$7;m,100000,$12,ss;p,0,@3,$13;m,ss,@3,100000;a,$6,$3,0;m,h,h,1;m,$15,$15,1;a,$17,0,$18;r,$19;m,$20,$6,h;a,$6,$21,0;a,$16,$21,0;a,$4,0,$3;m,$24,1,2;a,$22,0,$25;a,$23,$25,0;m,$4,1,2;a,r,0,$28;a,b,0,$15;t,$15,$21;a,0,0,$31;a,b,0,$12;d,$6,$16,2;m,$3,1,2;t,$12,$35;a,k34,$31,0;a,k34,0,$36;a,$36,0,cd4;a,cd4,$36,0|f,0,0,0:M,$29,b;L,$26,$30;L,$22,$30;A,$6,h,$37,$32;L,$16,t;A,$6,h,k34,$31;L,$27,$30;Z~f,0,0,0:M,$34,$33;A,$6,h,$38,$39;L,l,b;A,$6,h,cd2,$40;Z~n,1,0,0:M,$34,$33;A,$6,h,$38,$39;L,l,b;A,$6,h,cd2,cd4;L,$16,t;A,$6,h,k34,$31;L,$27,$30;L,$29,b;L,$26,$30;L,$22,$30;A,$6,h,$37,$32|l,t,r,b"
        Case 49      ' stripedRightArrow
            s = "50000,50000|m,84375,w,ss;p,0,@1,100000;p,0,@2,$0;m,ss,5,32;m,ss,$2,100000;a,r,0,$4;m,h,$1,200000;a,vc,0,$6;a,vc,$6,0;m,$6,$4,hd2;a,r,0,$9|f,1,0,0:M,l,$7;L,ssd32,$7;L,ssd32,$8;L,l,$8;Z;M,ssd16,$7;L,ssd8,$7;L,ssd8,$8;L,ssd16,$8;Z;M,$3,$7;L,$5,$7;L,$5,t;L,r,vc;L,$5,b;L,$5,$8;L,$3,$8;Z|$3,$7,$10,$8"
        Case 50      ' notchedRightArrow
            s = "50000,50000|m,100000,w,ss;p,0,@1,100000;p,0,@2,$0;m,ss,$2,100000;a,r,0,$3;m,h,$1,200000;a,vc,0,$5;a,vc,$5,0;m,$5,$3,hd2;a,r,0,$8|f,1,0,0:M,l,$6;L,$4,$6;L,$4,t;L,r,vc;L,$4,b;L,$4,$7;L,l,$7;L,$8,vc;Z|$8,$6,$9,$7"
        Case 51      ' homePlate
            s = "50000|m,100000,w,ss;p,0,@1,$0;m,ss,$1,100000;a,r,0,$2;d,$3,r,2;m,$3,1,2|f,1,0,0:M,l,t;L,$3,t;L,r,vc;L,$3,b;L,l,b;Z|l,t,$4,b"
        Case 52      ' chevron
            s = "50000|m,100000,w,ss;p,0,@1,$0;m,ss,$1,100000;a,r,0,$2;m,$3,1,2;a,$3,0,$2;q,$5,$2,l;q,$5,$3,r|f,1,0,0:M,l,t;L,$3,t;L,r,vc;L,$3,b;L,l,b;L,$2,vc;Z|$6,t,$7,b"
        Case 53      ' rightArrowCallout
            s = "25000,25000,25000,64977|m,50000,h,ss;p,0,@2,$0;m,$1,2,1;p,0,@1,$2;m,100000,w,ss;p,0,@3,$4;m,$5,ss,w;a,100000,0,$6;p,0,@4,$7;m,ss,$1,100000;m,ss,$3,200000;a,vc,0,$9;a,vc,0,$10;a,vc,$10,0;a,vc,$9,0;m,ss,$5,100000;a,r,0,$15;m,w,$8,100000;m,$17,1,2|f,1,0,0:M,l,t;L,$17,t;L,$17,$12;L,$16,$12;L,$16,$11;L,r,vc;L,$16,$14;L,$16,$13;L,$17,$13;L,$17,b;L,l,b;Z|l,t,$17,b"
        Case 54      ' leftArrowCallout
            s = "25000,25000,25000,64977|m,50000,h,ss;p,0,@2,$0;m,$1,2,1;p,0,@1,$2;m,100000,w,ss;p,0,@3,$4;m,$5,ss,w;a,100000,0,$6;p,0,@4,$7;m,ss,$1,100000;m,ss,$3,200000;a,vc,0,$9;a,vc,0,$10;a,vc,$10,0;a,vc,$9,0;m,ss,$5,100000;m,w,$8,100000;a,r,0,$16;d,$17,r,2|f,1,0,0:M,l,vc;L,$15,$11;L,$15,$12;L,$17,$12;L,$17,t;L,r,t;L,r,b;L,$17,b;L,$17,$13;L,$15,$13;L,$15,$14;Z|$17,t,r,b"
        Case 55      ' upArrowCallout
            s = "25000,25000,25000,64977|m,50000,w,ss;p,0,@2,$0;m,$1,2,1;p,0,@1,$2;m,100000,h,ss;p,0,@3,$4;m,$5,ss,h;a,100000,0,$6;p,0,@4,$7;m,ss,$1,100000;m,ss,$3,200000;a,hc,0,$9;a,hc,0,$10;a,hc,$10,0;a,hc,$9,0;m,ss,$5,100000;m,h,$8,100000;a,b,0,$16;d,$17,b,2|f,1,0,0:M,l,$17;L,$12,$17;L,$12,$15;L,$11,$15;L,hc,t;L,$14,$15;L,$13,$15;L,$13,$17;L,r,$17;L,r,b;L,l,b;Z|l,$17,r,b"
        Case 56      ' downArrowCallout
            s = "25000,25000,25000,64977|m,50000,w,ss;p,0,@2,$0;m,$1,2,1;p,0,@1,$2;m,100000,h,ss;p,0,@3,$4;m,$5,ss,h;a,100000,0,$6;p,0,@4,$7;m,ss,$1,100000;m,ss,$3,200000;a,hc,0,$9;a,hc,0,$10;a,hc,$10,0;a,hc,$9,0;m,ss,$5,100000;a,b,0,$15;m,h,$8,100000;m,$17,1,2|f,1,0,0:M,l,t;L,r,t;L,r,$17;L,$13,$17;L,$13,$16;L,$14,$16;L,hc,b;L,$11,$16;L,$12,$16;L,$12,$17;L,l,$17;Z|l,t,r,$17"
        Case 57      ' leftRightArrowCallout
            s = "25000,25000,25000,48123|m,50000,h,ss;p,0,@2,$0;m,$1,2,1;p,0,@1,$2;m,50000,w,ss;p,0,@3,$4;m,$5,ss,wd2;a,100000,0,$6;p,0,@4,$7;m,ss,$1,100000;m,ss,$3,200000;a,vc,0,$9;a,vc,0,$10;a,vc,$10,0;a,vc,$9,0;m,ss,$5,100000;a,r,0,$15;m,w,$8,200000;a,hc,0,$17;a,hc,$17,0|f,1,0,0:M,l,vc;L,$15,$11;L,$15,$12;L,$18,$12;L,$18,t;L,$19,t;L,$19,$12;L,$16,$12;L,$16,$11;L,r,vc;L,$16,$14;L,$16,$13;L,$19,$13;L,$19,b;L,$18,b;L,$18,$13;L,$15,$13;L,$15,$14;Z|$18,t,$19,b"
        Case 58      ' upDownArrowCallout
            s = "25000,25000,25000,48123|m,50000,w,ss;p,0,@2,$0;m,$1,2,1;p,0,@1,$2;m,50000,h,ss;p,0,@3,$4;m,$5,ss,hd2;a,100000,0,$6;p,0,@4,$7;m,ss,$1,100000;m,ss,$3,200000;a,hc,0,$9;a,hc,0,$10;a,hc,$10,0;a,hc,$9,0;m,ss,$5,100000;a,b,0,$15;m,h,$8,200000;a,vc,0,$17;a,vc,$17,0|f,1,0,0:M,l,$18;L,$12,$18;L,$12,$15;L,$11,$15;L,hc,t;L,$14,$15;L,$13,$15;L,$13,$18;L,r,$18;L,r,$19;L,$13,$19;L,$13,$16;L,$14,$16;L,hc,b;L,$11,$16;L,$12,$16;L,$12,$19;L,l,$19;Z|l,$18,r,$19"
        Case 59      ' quadArrowCallout
            s = "18515,18515,18515,48123|p,0,@2,50000;m,$0,2,1;p,0,@1,$1;a,50000,0,$0;p,0,@3,$3;m,$4,2,1;a,100000,0,$5;p,$2,@4,$6;m,ss,$0,100000;m,ss,$2,200000;m,ss,$4,100000;m,w,$7,200000;m,h,$7,200000;a,r,0,$10;a,hc,0,$11;a,hc,$11,0;a,hc,0,$8;a,hc,$8,0;a,hc,0,$9;a,hc,$9,0;a,b,0,$10;a,vc,0,$12;a,vc,$12,0;a,vc,0,$8;a,vc,$8,0;a,vc,0,$9;a,vc,$9,0|f,1,0,0:M,l,vc;L,$10,$23;L,$10,$25;L,$14,$25;L,$14,$21;L,$18,$21;L,$18,$10;L,$16,$10;L,hc,t;L,$17,$10;L,$19,$10;L,$19,$21;L,$15,$21;L,$15,$25;L,$13,$25;L,$13,$23;L,r,vc;L,$13,$24;L,$13,$26;L,$15,$26;L,$15,$22;L,$19,$22;L,$19,$20;L,$17,$20;L,hc,b;L,$16,$20;L,$18,$20;L,$18,$22;L,$14,$22;L,$14,$26;L,$10,$26;L,$10,$24;Z|$14,$21,$15,$22"
        Case 60      ' circularArrow
            s = "12500,!1142319,!20457681,!10800000,12500|p,0,@5,25000;m,$0,2,1;p,0,@1,$1;p,1,@3,21599999;p,0,@4,21599999;m,ss,$2,100000;m,ss,$0,100000;m,$5,1,2;a,wd2,$7,$6;a,hd2,$7,$6;a,$8,0,$5;a,$9,0,$5;a,$10,$7,0;a,$11,$7,0;i,$12,$3;o,$13,$3;c,$12,$15,$14;s,$13,$15,$14;a,hc,$16,0;a,vc,$17,0;n,$10,$11;m,$16,$16,1;m,$17,$17,1;m,$20,$20,1;a,$21,0,$23;a,$22,0,$23;m,$24,$25,$21;m,$26,1,$22;a,1,0,$27;r,$28;m,$24,1,$16;m,$30,1,$17;d,1,$29,$31;t,1,$32;a,$33,21600000,0;q,$33,$33,$34;a,$35,0,$3;a,$36,21600000,0;q,$36,$36,$37;a,$38,0,cd2;a,$38,0,21600000;q,$39,$40,$38;b,$41;p,0,@2,$42;a,$3,$43,0;i,$12,$44;o,$13,$44;c,$12,$46,$45;s,$13,$46,$45;a,hc,$47,0;a,vc,$48,0;i,$8,$4;o,$9,$4;c,$8,$52,$51;s,$9,$52,$51;a,hc,$53,0;a,vc,$54,0;o,$6,$44;i,$6,$44;a,$18,$57,0;a,$19,$58,0;o,$6,$44;i,$6,$44;a,$18,0,$61,0;a,$19,0,$62,0;a,$63,0,hc;a,$64,0,vc;a,$59,0,hc;a,$60,0,vc;n,$8,$9;m,$65,$69,$8;m,$66,$69,$9;m,$67,$69,$8;m,$68,$69"
            s = s & ",$9;a,$72,0,$70;a,$73,0,$71;h,$74,$75,0;m,$70,$73,1;m,$72,$71,1;a,$77,0,$78;m,$69,$69,1;m,$76,$76,1;m,$80,$81,1;m,$79,$79,1;a,$82,0,$83;x,$84,0;r,$85;m,$75,-1,1;q,$87,-1,1;m,$88,$74,1;m,$89,$86,1;m,$79,$75,1;d,$91,$90,$81;a,$91,0,$90;m,$93,1,$81;b,$75;m,$95,$86,1;m,$79,$74,-1;d,$97,$96,$81;a,$97,0,$96;m,$99,1,$81;a,$72,0,$92;a,$72,0,$94;a,$73,0,$98;a,$73,0,$100;h,$101,$103,0;h,$102,$104,0;a,$106,0,$105;q,$107,$92,$94;q,$107,$98,$100;m,$108,$8,$69;m,$109,$9,$69;a,hc,$110,0;a,vc,$111,0;m,$65,$20,$10;m,$66,$20,$11;m,$67,$20,$10;m,$68,$20,$11;a,$116,0,$114;a,$117,0,$115;h,$118,$119,0;m,$114,$117,1;m,$116,$115,1;a,$121,0,$122;m,$20,$20,1;m,$120,$120,1;m,$124,$125,1;m,$123,$123,1;a,$126,0,$127;x,$128,0;r,$129;m,$88,$118,1;m,$131,$130,1;m,$123,$119,1;d,$133,$132,$125;a,$133,0,$132;m,$135,1,$125;b,$119;m,$137,$130,1;m,$123,$118,-1;d,$139,$138,$125;a,$139,0,$138;m,$141,1,$125;a,$114,0,$134;a,$114"
            s = s & ",0,$136;a,$115,0,$140;a,$115,0,$142;h,$143,$145,0;h,$144,$146,0;a,$148,0,$147;q,$149,$134,$136;q,$149,$140,$142;m,$150,$10,$20;m,$151,$11,$20;a,hc,$152,0;a,vc,$153,0;t,$152,$153;a,$156,21600000,0;q,$156,$156,$157;a,$4,0,$158;a,$159,0,21600000;q,$159,$160,$159;a,$112,0,$154;a,$113,0,$155;h,$162,$163,0;m,$164,1,2;a,$165,0,$6;q,$166,$112,$59;q,$166,$113,$60;q,$166,$154,$63;q,$166,$155,$64;t,$110,$111;a,$171,21600000,0;q,$171,$171,$172;a,$173,0,$4;a,$174,21600000,0;q,$174,$174,$175;i,$12,$4;o,$13,$4;c,$12,$178,$177;s,$13,$178,$177;a,hc,$179,0;a,vc,$180,0;a,$4,0,cd4;a,$44,cd4,0;a,$44,cd2,0;o,$8,2700000;i,$9,2700000;a,hc,0,$186;a,hc,$186,0;a,vc,0,$187;a,vc,$187,0|f,1,0,0:M,$55,$56;A,$8,$9,$4,$176;L,$167,$168;L,$49,$50;L,$169,$170;L,$154,$155;A,$10,$11,$158,$161;Z|$188,$190,$189,$191"
        Case 61      ' flowChartProcess
            s = "||f,1,1,1:M,0,0;L,1,0;L,1,1;L,0,1;Z|l,t,r,b"
        Case 62      ' flowChartAlternateProcess
            s = "|a,r,0,ssd6;a,b,0,ssd6;m,ssd6,29289,100000;a,r,0,$2;a,b,0,$2|f,1,0,0:M,l,ssd6;A,ssd6,ssd6,cd2,cd4;L,$0,t;A,ssd6,ssd6,k34,cd4;L,r,$1;A,ssd6,ssd6,0,cd4;L,ssd6,b;A,ssd6,ssd6,cd4,cd4;Z|$2,$2,$3,$4"
        Case 63      ' flowChartDecision
            s = "|m,w,3,4;m,h,3,4|f,1,2,2:M,0,1;L,1,0;L,2,1;L,1,2;Z|wd4,hd4,$0,$1"
        Case 64      ' flowChartInputOutput
            s = "|m,w,2,5;m,w,3,5;m,w,4,5;m,w,9,10|f,1,5,5:M,0,5;L,1,0;L,5,0;L,4,5;Z|wd5,t,$2,b"
        Case 65      ' flowChartPredefinedProcess
            s = "|m,w,7,8|f,0,1,1:M,0,0;L,1,0;L,1,1;L,0,1;Z~n,1,8,8:M,1,0;L,1,8;M,7,0;L,7,8~n,1,1,1:M,0,0;L,1,0;L,1,1;L,0,1;Z|wd8,t,$0,b"
        Case 66      ' flowChartInternalStorage
            s = "||f,0,1,1:M,0,0;L,1,0;L,1,1;L,0,1;Z~n,1,8,8:M,1,0;L,1,8;M,0,1;L,8,1~n,1,1,1:M,0,0;L,1,0;L,1,1;L,0,1;Z|wd8,hd8,r,b"
        Case 67      ' flowChartDocument
            s = "|m,h,17322,21600;m,h,20172,21600|f,1,21600,21600:M,0,0;L,21600,0;L,21600,17322;C,10800,17322,10800,23922,0,20172;Z|l,t,r,$0"
        Case 68      ' flowChartMultidocument
            s = "|m,h,3675,21600;m,h,20782,21600;m,w,9298,21600;m,w,12286,21600;m,w,18595,21600|f,0,21600,21600:M,0,20782;C,9298,23542,9298,18022,18595,18022;L,18595,3675;L,0,3675;Z;M,1532,3675;L,1532,1815;L,20000,1815;L,20000,16252;C,19298,16252,18595,16352,18595,16352;L,18595,3675;Z;M,2972,1815;L,2972,0;L,21600,0;L,21600,14392;C,20800,14392,20000,14467,20000,14467;L,20000,1815;Z~n,1,21600,21600:M,0,3675;L,18595,3675;L,18595,18022;C,9298,18022,9298,23542,0,20782;Z;M,1532,3675;L,1532,1815;L,20000,1815;L,20000,16252;C,19298,16252,18595,16352,18595,16352;M,2972,1815;L,2972,0;L,21600,0;L,21600,14392;C,20800,14392,20000,14467,20000,14467~n,0,21600,21600:M,0,20782;C,9298,23542,9298,18022,18595,18022;L,18595,16352;C,18595,16352,19298,16252,20000,16252;L,20000,14467;C,20000,14467,20800,14392,21600,14392;L,21600,0;L,2972,0;L,2972,1815;L,1532,1815;L,1532,3675;L,0,3675;Z|l,$0,$4,$1"
        Case 69      ' flowChartTerminator
            s = "|m,w,1018,21600;m,w,20582,21600;m,h,3163,21600;m,h,18437,21600|f,1,21600,21600:M,3475,0;L,18125,0;A,3475,10800,k34,cd2;L,3475,21600;A,3475,10800,cd4,cd2;Z|$0,$2,$1,$3"
        Case 70      ' flowChartPreparation
            s = "|m,w,4,5|f,1,10,10:M,0,5;L,2,0;L,8,0;L,10,5;L,8,10;L,2,10;Z|wd5,t,$0,b"
        Case 71      ' flowChartManualInput
            s = "||f,1,5,5:M,0,1;L,5,0;L,5,5;L,0,5;Z|l,hd5,r,b"
        Case 72      ' flowChartManualOperation
            s = "|m,w,4,5;m,w,9,10|f,1,5,5:M,0,0;L,5,0;L,4,5;L,1,5;Z|wd5,t,$0,b"
        Case 73      ' flowChartConnector
            s = "|o,wd2,2700000;i,hd2,2700000;a,hc,0,$0;a,hc,$0,0;a,vc,0,$1;a,vc,$1,0|f,1,0,0:M,l,vc;A,wd2,hd2,cd2,cd4;A,wd2,hd2,k34,cd4;A,wd2,hd2,0,cd4;A,wd2,hd2,cd4,cd4;Z|$2,$4,$3,$5"
        Case 74      ' flowChartOffpageConnector
            s = "|m,h,4,5|f,1,10,10:M,0,0;L,10,0;L,10,8;L,5,10;L,0,8;Z|l,t,r,$0"
        Case 75      ' flowChartPunchedCard
            s = "||f,1,5,5:M,0,1;L,1,0;L,5,0;L,5,5;L,0,5;Z|l,hd5,r,b"
        Case 76      ' flowChartPunchedTape
            s = "|m,h,9,10;m,h,4,5|f,1,20,20:M,0,2;A,5,2,cd2,-10800000;A,5,2,cd2,cd2;L,20,18;A,5,2,0,-10800000;A,5,2,0,cd2;Z|l,hd5,r,$1"
    End Select
    PresetSrc1 = s
End Function

Private Function PresetSrc2(ByVal ast As Long) As String
    Dim s As String
    Select Case ast
        Case 77      ' flowChartSummingJunction
            s = "|o,wd2,2700000;i,hd2,2700000;a,hc,0,$0;a,hc,$0,0;a,vc,0,$1;a,vc,$1,0|f,0,0,0:M,l,vc;A,wd2,hd2,cd2,cd4;A,wd2,hd2,k34,cd4;A,wd2,hd2,0,cd4;A,wd2,hd2,cd4,cd4;Z~n,1,0,0:M,$2,$4;L,$3,$5;M,$3,$4;L,$2,$5~n,1,0,0:M,l,vc;A,wd2,hd2,cd2,cd4;A,wd2,hd2,k34,cd4;A,wd2,hd2,0,cd4;A,wd2,hd2,cd4,cd4;Z|$2,$4,$3,$5"
        Case 78      ' flowChartOr
            s = "|o,wd2,2700000;i,hd2,2700000;a,hc,0,$0;a,hc,$0,0;a,vc,0,$1;a,vc,$1,0|f,0,0,0:M,l,vc;A,wd2,hd2,cd2,cd4;A,wd2,hd2,k34,cd4;A,wd2,hd2,0,cd4;A,wd2,hd2,cd4,cd4;Z~n,1,0,0:M,hc,t;L,hc,b;M,l,vc;L,r,vc~n,1,0,0:M,l,vc;A,wd2,hd2,cd2,cd4;A,wd2,hd2,k34,cd4;A,wd2,hd2,0,cd4;A,wd2,hd2,cd4,cd4;Z|$2,$4,$3,$5"
        Case 79      ' flowChartCollate
            s = "|m,w,3,4;m,h,3,4|f,1,2,2:M,0,0;L,2,0;L,1,1;L,2,2;L,0,2;L,1,1;Z|wd4,hd4,$0,$1"
        Case 80      ' flowChartSort
            s = "|m,w,3,4;m,h,3,4|f,0,2,2:M,0,1;L,1,0;L,2,1;L,1,2;Z~n,1,2,2:M,0,1;L,2,1~n,1,2,2:M,0,1;L,1,0;L,2,1;L,1,2;Z|wd4,hd4,$0,$1"
        Case 81      ' flowChartExtract
            s = "|m,w,3,4|f,1,2,2:M,0,2;L,1,0;L,2,2;Z|wd4,vc,$0,b"
        Case 82      ' flowChartMerge
            s = "|m,w,3,4|f,1,2,2:M,0,0;L,2,0;L,1,2;Z|wd4,t,$0,vc"
        Case 83      ' flowChartOnlineStorage
            s = "|m,w,5,6|f,1,6,6:M,1,0;L,6,0;A,1,3,k34,-10800000;L,1,6;A,1,3,cd4,cd2;Z|wd6,t,$0,b"
        Case 84      ' flowChartDelay
            s = "|o,wd2,2700000;i,hd2,2700000;a,hc,$0,0;a,vc,0,$1;a,vc,$1,0|f,1,0,0:M,l,t;L,hc,t;A,wd2,hd2,k34,cd2;L,l,b;Z|l,$3,$2,$4"
        Case 85      ' flowChartMagneticTape
            s = "|o,wd2,2700000;i,hd2,2700000;a,hc,0,$0;a,hc,$0,0;a,vc,0,$1;a,vc,$1,0;t,w,h|f,1,0,0:M,hc,b;A,wd2,hd2,cd4,cd4;A,wd2,hd2,cd2,cd4;A,wd2,hd2,k34,cd4;A,wd2,hd2,0,$6;L,r,$5;L,r,b;Z|$2,$4,$3,$5"
        Case 86      ' flowChartMagneticDisk
            s = "|m,h,5,6|f,0,6,6:M,0,1;A,3,1,cd2,cd2;L,6,5;A,3,1,0,cd2;Z~n,1,6,6:M,6,1;A,3,1,0,cd2~n,1,6,6:M,0,1;A,3,1,cd2,cd2;L,6,5;A,3,1,0,cd2;Z|l,hd3,r,$0"
        Case 87      ' flowChartMagneticDrum
            s = "|m,w,2,3|f,0,6,6:M,1,0;L,5,0;A,1,3,k34,cd2;L,1,6;A,1,3,cd4,cd2;Z~n,1,6,6:M,5,6;A,1,3,cd4,cd2~n,1,6,6:M,1,0;L,5,0;A,1,3,k34,cd2;L,1,6;A,1,3,cd4,cd2;Z|wd6,t,$0,b"
        Case 88      ' flowChartDisplay
            s = "|m,w,5,6|f,1,6,6:M,0,3;L,1,0;L,5,0;A,1,3,k34,cd2;L,1,6;Z|wd6,t,$0,b"
        Case 89      ' irregularSeal1
            s = "|m,w,4627,21600;m,w,8485,21600;m,w,16702,21600;m,w,14522,21600;m,h,6320,21600;m,h,8615,21600;m,h,13937,21600;m,h,13290,21600|f,1,21600,21600:M,10800,5800;L,14522,0;L,14155,5325;L,18380,4457;L,16702,7315;L,21097,8137;L,17607,10475;L,21600,13290;L,16837,12942;L,18145,18095;L,14020,14457;L,13247,19737;L,10532,14935;L,8485,21600;L,7715,15627;L,4762,17617;L,5667,13937;L,135,14587;L,3722,11775;L,0,8615;L,4627,7617;L,370,2295;L,7312,6320;L,8352,2295;Z|$0,$4,$2,$6"
        Case 90      ' irregularSeal2
            s = "|m,w,9722,21600;m,w,5372,21600;m,w,11612,21600;m,w,14640,21600;m,h,1887,21600;m,h,6382,21600;m,h,12877,21600;m,h,19712,21600;m,h,18842,21600;m,h,15935,21600;m,h,6645,21600|f,1,21600,21600:M,11462,4342;L,14790,0;L,14525,5777;L,18007,3172;L,16380,6532;L,21600,6645;L,16985,9402;L,18270,11290;L,16380,12310;L,18877,15632;L,14640,14350;L,14942,17370;L,12180,15935;L,11612,18842;L,9872,17370;L,8700,19712;L,7527,18125;L,4917,21600;L,4805,18240;L,1285,17825;L,3330,15370;L,0,12877;L,3935,11592;L,1172,8270;L,5372,7817;L,4502,3625;L,8550,6382;L,9722,1887;Z|$1,$5,$3,$9"
        Case 91      ' star4
            s = "12500|p,0,@1,50000;m,wd2,$0,50000;m,hd2,$0,50000;o,$1,2700000;i,$2,2700000;a,hc,0,$3;a,hc,$3,0;a,vc,0,$4;a,vc,$4,0;a,vc,0,$2|f,1,0,0:M,l,vc;L,$5,$7;L,hc,t;L,$6,$7;L,r,vc;L,$6,$8;L,hc,b;L,$5,$8;Z|$5,$7,$6,$8"
        Case 92      ' star5
            s = "19098,105146,110557|p,0,@1,50000;m,wd2,@2,100000;m,hd2,@3,100000;m,vc,@3,100000;o,$1,1080000;o,$1,18360000;i,$2,1080000;i,$2,18360000;a,hc,0,$4;a,hc,0,$5;a,hc,$5,0;a,hc,$4,0;a,$3,0,$6;a,$3,0,$7;m,$1,$0,50000;m,$2,$0,50000;o,$14,20520000;o,$14,3240000;i,$15,3240000;i,$15,20520000;a,hc,0,$16;a,hc,0,$17;a,hc,$17,0;a,hc,$16,0;a,$3,0,$18;a,$3,0,$19;a,$3,$15,0;a,$3,0,$15|f,1,0,0:M,$8,$12;L,$21,$24;L,hc,t;L,$22,$24;L,$11,$12;L,$23,$25;L,$10,$13;L,hc,$26;L,$9,$13;L,$20,$25;Z|$20,$24,$23,$26"
        Case 93      ' star8
            s = "37500|p,0,@1,50000;o,wd2,2700000;a,hc,0,$1;a,hc,$1,0;i,hd2,2700000;a,vc,0,$4;a,vc,$4,0;m,wd2,$0,50000;m,hd2,$0,50000;m,$7,92388,100000;m,$7,38268,100000;m,$8,92388,100000;m,$8,38268,100000;a,hc,0,$9;a,hc,0,$10;a,hc,$10,0;a,hc,$9,0;a,vc,0,$11;a,vc,0,$12;a,vc,$12,0;a,vc,$11,0;a,vc,0,$8|f,1,0,0:M,l,vc;L,$13,$18;L,$2,$5;L,$14,$17;L,hc,t;L,$15,$17;L,$3,$5;L,$16,$18;L,r,vc;L,$16,$19;L,$3,$6;L,$15,$20;L,hc,b;L,$14,$20;L,$2,$6;L,$13,$19;Z|$13,$17,$16,$20"
        Case 94      ' star16
            s = "37500|p,0,@1,50000;m,wd2,92388,100000;m,wd2,70711,100000;m,wd2,38268,100000;m,hd2,92388,100000;m,hd2,70711,100000;m,hd2,38268,100000;a,hc,0,$1;a,hc,0,$2;a,hc,0,$3;a,hc,$3,0;a,hc,$2,0;a,hc,$1,0;a,vc,0,$4;a,vc,0,$5;a,vc,0,$6;a,vc,$6,0;a,vc,$5,0;a,vc,$4,0;m,wd2,$0,50000;m,hd2,$0,50000;m,$19,98079,100000;m,$19,83147,100000;m,$19,55557,100000;m,$19,19509,100000;m,$20,98079,100000;m,$20,83147,100000;m,$20,55557,100000;m,$20,19509,100000;a,hc,0,$21;a,hc,0,$22;a,hc,0,$23;a,hc,0,$24;a,hc,$24,0;a,hc,$23,0;a,hc,$22,0;a,hc,$21,0;a,vc,0,$25;a,vc,0,$26;a,vc,0,$27;a,vc,0,$28;a,vc,$28,0;a,vc,$27,0;a,vc,$26,0;a,vc,$25,0;o,$19,2700000;i,$20,2700000;a,hc,0,$45;a,vc,0,$46;a,hc,$45,0;a,vc,$46,0;a,vc,0,$20|f,1,0,0:M,l,vc;L,$29,$40;L,$7,$15;L,$30,$39;L,$8,$14;L,$31,$38;L,$9,$13;L,$32,$37;L,hc,t;L,$33,$37;L,$10,$13;L,$34,$38;L,$11,$14;L,$35,$39;L,$12,$15;L,$36,$40;L,r,vc;L,$36,$41;L,$12,$16;L,$35,$42;L,$11,$17;"
            s = s & "L,$34,$43;L,$10,$18;L,$33,$44;L,hc,b;L,$32,$44;L,$9,$18;L,$31,$43;L,$8,$17;L,$30,$42;L,$7,$16;L,$29,$41;Z|$47,$48,$49,$50"
        Case 95      ' star24
            s = "37500|p,0,@1,50000;o,wd2,900000;o,wd2,1800000;o,wd2,2700000;v,wd4;o,wd2,4500000;i,hd2,4500000;i,hd2,3600000;i,hd2,2700000;v,hd4;i,hd2,900000;a,hc,0,$1;a,hc,0,$2;a,hc,0,$3;a,hc,0,$4;a,hc,0,$5;a,hc,$5,0;a,hc,$4,0;a,hc,$3,0;a,hc,$2,0;a,hc,$1,0;a,vc,0,$6;a,vc,0,$7;a,vc,0,$8;a,vc,0,$9;a,vc,0,$10;a,vc,$10,0;a,vc,$9,0;a,vc,$8,0;a,vc,$7,0;a,vc,$6,0;m,wd2,$0,50000;m,hd2,$0,50000;m,$31,99144,100000;m,$31,92388,100000;m,$31,79335,100000;m,$31,60876,100000;m,$31,38268,100000;m,$31,13053,100000;m,$32,99144,100000;m,$32,92388,100000;m,$32,79335,100000;m,$32,60876,100000;m,$32,38268,100000;m,$32,13053,100000;a,hc,0,$33;a,hc,0,$34;a,hc,0,$35;a,hc,0,$36;a,hc,0,$37;a,hc,0,$38;a,hc,$38,0;a,hc,$37,0;a,hc,$36,0;a,hc,$35,0;a,hc,$34,0;a,hc,$33,0;a,vc,0,$39;a,vc,0,$40;a,vc,0,$41;a,vc,0,$42;a,vc,0,$43;a,vc,0,$44;a,vc,$44,0;a,vc,$43,0;a,vc,$42,0;a,vc,$41,0;a,vc,$40,0;a,vc,$39,0;o,$31,2700000;i,$32,2700000;a,hc,0,"
            s = s & "$69;a,vc,0,$70;a,hc,$69,0;a,vc,$70,0;a,vc,0,$32|f,1,0,0:M,l,vc;L,$45,$62;L,$11,$25;L,$46,$61;L,$12,$24;L,$47,$60;L,$13,$23;L,$48,$59;L,$14,$22;L,$49,$58;L,$15,$21;L,$50,$57;L,hc,t;L,$51,$57;L,$16,$21;L,$52,$58;L,$17,$22;L,$53,$59;L,$18,$23;L,$54,$60;L,$19,$24;L,$55,$61;L,$20,$25;L,$56,$62;L,r,vc;L,$56,$63;L,$20,$26;L,$55,$64;L,$19,$27;L,$54,$65;L,$18,$28;L,$53,$66;L,$17,$29;L,$52,$67;L,$16,$30;L,$51,$68;L,hc,b;L,$50,$68;L,$15,$30;L,$49,$67;L,$14,$29;L,$48,$66;L,$13,$28;L,$47,$65;L,$12,$27;L,$46,$64;L,$11,$26;L,$45,$63;Z|$71,$72,$73,$74"
        Case 96      ' star32
            s = "37500|p,0,@1,50000;m,wd2,98079,100000;m,wd2,92388,100000;m,wd2,83147,100000;o,wd2,2700000;m,wd2,55557,100000;m,wd2,38268,100000;m,wd2,19509,100000;m,hd2,98079,100000;m,hd2,92388,100000;m,hd2,83147,100000;i,hd2,2700000;m,hd2,55557,100000;m,hd2,38268,100000;m,hd2,19509,100000;a,hc,0,$1;a,hc,0,$2;a,hc,0,$3;a,hc,0,$4;a,hc,0,$5;a,hc,0,$6;a,hc,0,$7;a,hc,$7,0;a,hc,$6,0;a,hc,$5,0;a,hc,$4,0;a,hc,$3,0;a,hc,$2,0;a,hc,$1,0;a,vc,0,$8;a,vc,0,$9;a,vc,0,$10;a,vc,0,$11;a,vc,0,$12;a,vc,0,$13;a,vc,0,$14;a,vc,$14,0;a,vc,$13,0;a,vc,$12,0;a,vc,$11,0;a,vc,$10,0;a,vc,$9,0;a,vc,$8,0;m,wd2,$0,50000;m,hd2,$0,50000;m,$43,99518,100000;m,$43,95694,100000;m,$43,88192,100000;m,$43,77301,100000;m,$43,63439,100000;m,$43,47140,100000;m,$43,29028,100000;m,$43,9802,100000;m,$44,99518,100000;m,$44,95694,100000;m,$44,88192,100000;m,$44,77301,100000;m,$44,63439,100000;m,$44,47140,100000;m,$44,29028,100000;m,$44,9802,100000;a,h"
            s = s & "c,0,$45;a,hc,0,$46;a,hc,0,$47;a,hc,0,$48;a,hc,0,$49;a,hc,0,$50;a,hc,0,$51;a,hc,0,$52;a,hc,$52,0;a,hc,$51,0;a,hc,$50,0;a,hc,$49,0;a,hc,$48,0;a,hc,$47,0;a,hc,$46,0;a,hc,$45,0;a,vc,0,$53;a,vc,0,$54;a,vc,0,$55;a,vc,0,$56;a,vc,0,$57;a,vc,0,$58;a,vc,0,$59;a,vc,0,$60;a,vc,$60,0;a,vc,$59,0;a,vc,$58,0;a,vc,$57,0;a,vc,$56,0;a,vc,$55,0;a,vc,$54,0;a,vc,$53,0;o,$43,2700000;i,$44,2700000;a,hc,0,$93;a,vc,0,$94;a,hc,$93,0;a,vc,$94,0;a,vc,0,$44|f,1,0,0:M,l,vc;L,$61,$84;L,$15,$35;L,$62,$83;L,$16,$34;L,$63,$82;L,$17,$33;L,$64,$81;L,$18,$32;L,$65,$80;L,$19,$31;L,$66,$79;L,$20,$30;L,$67,$78;L,$21,$29;L,$68,$77;L,hc,t;L,$69,$77;L,$22,$29;L,$70,$78;L,$23,$30;L,$71,$79;L,$24,$31;L,$72,$80;L,$25,$32;L,$73,$81;L,$26,$33;L,$74,$82;L,$27,$34;L,$75,$83;L,$28,$35;L,$76,$84;L,r,vc;L,$76,$85;L,$28,$36;L,$75,$86;L,$27,$37;L,$74,$87;L,$26,$38;L,$73,$88;L,$25,$39;L,$72,$89;L,$24,$40;L,$71,$90;L,$23,$41;L,$70,$91;L,$22,$42"
            s = s & ";L,$69,$92;L,hc,b;L,$68,$92;L,$21,$42;L,$67,$91;L,$20,$41;L,$66,$90;L,$19,$40;L,$65,$89;L,$18,$39;L,$64,$88;L,$17,$38;L,$63,$87;L,$16,$37;L,$62,$86;L,$15,$36;L,$61,$85;Z|$95,$96,$97,$98"
        Case 97      ' ribbon2
            s = "16667,50000|p,0,@1,33333;p,25000,@2,75000;a,r,0,wd8;m,w,$1,200000;a,hc,0,$3;a,hc,$3,0;a,$4,wd32,0;a,$5,0,wd32;a,$4,wd8,0;a,$5,0,wd8;a,$8,0,wd32;a,$9,wd32,0;m,h,$0,200000;a,b,0,$12;m,h,$0,100000;a,b,0,$14;a,t,$14,0;d,$16,b,2;m,h,$0,400000;a,b,0,$18;a,$13,0,$18|f,0,0,0:M,l,b;L,$10,b;A,wd32,$18,cd4,-10800000;L,$6,$13;A,wd32,$18,cd4,cd2;L,$7,$15;A,wd32,$18,k34,cd2;L,$11,$13;A,wd32,$18,k34,-10800000;L,r,b;L,$2,$17;L,r,$16;L,$5,$16;L,$5,$18;A,wd32,$18,0,-5400000;L,$6,t;A,wd32,$18,k34,-5400000;L,$4,$16;L,l,$16;L,wd8,$17;Z~f,0,0,0:M,$8,$19;A,wd32,$18,0,-5400000;L,$6,$13;A,wd32,$18,cd4,cd2;L,$8,$15;Z;M,$9,$19;A,wd32,$18,cd2,cd4;L,$7,$13;A,wd32,$18,cd4,-10800000;L,$9,$15;Z~n,1,0,0:M,l,b;L,wd8,$17;L,l,$16;L,$4,$16;L,$4,$18;A,wd32,$18,cd2,cd4;L,$7,t;A,wd32,$18,k34,cd4;L,$5,$16;L,$5,$16;L,r,$16;L,$2,$17;L,r,b;L,$11,b;A,wd32,$18,cd4,cd2;L,$7,$13;A,wd32,$18,cd4,-10800000;L,$6,$15;A,wd32,$18,k34,-108000"
            s = s & "00;L,$10,$13;A,wd32,$18,k34,cd2;Z;M,$8,$15;L,$8,$19;M,$9,$19;L,$9,$15;M,$4,$20;L,$4,$16;M,$5,$16;L,$5,$20|$4,t,$5,$15"
        Case 98      ' ribbon
            s = "16667,50000|p,0,@1,33333;p,25000,@2,75000;a,r,0,wd8;m,w,$1,200000;a,hc,0,$3;a,hc,$3,0;a,$4,wd32,0;a,$5,0,wd32;a,$4,wd8,0;a,$5,0,wd8;a,$8,0,wd32;a,$9,wd32,0;m,h,$0,200000;m,h,$0,100000;a,b,0,$13;m,$14,1,2;m,h,$0,400000;a,b,0,$16;a,$13,0,$16|f,0,0,0:M,l,t;L,$10,t;A,wd32,$16,k34,cd2;L,$6,$12;A,wd32,$16,k34,-10800000;L,$7,$13;A,wd32,$16,cd4,-10800000;L,$11,$12;A,wd32,$16,cd4,cd2;L,r,t;L,$2,$15;L,r,$14;L,$5,$14;L,$5,$17;A,wd32,$16,0,cd4;L,$6,b;A,wd32,$16,cd4,cd4;L,$4,$14;L,l,$14;L,wd8,$15;Z~f,0,0,0:M,$8,$16;A,wd32,$16,0,cd4;L,$6,$12;A,wd32,$16,k34,-10800000;L,$8,$13;Z;M,$9,$16;A,wd32,$16,cd2,-5400000;L,$7,$12;A,wd32,$16,k34,cd2;L,$9,$13;Z~n,1,0,0:M,l,t;L,$10,t;A,wd32,$16,k34,cd2;L,$6,$12;A,wd32,$16,k34,-10800000;L,$7,$13;A,wd32,$16,cd4,-10800000;L,$11,$12;A,wd32,$16,cd4,cd2;L,r,t;L,$2,$15;L,r,$14;L,$5,$14;L,$5,$17;A,wd32,$16,0,cd4;L,$6,b;A,wd32,$16,cd4,cd4;L,$4,$14;L,l,$14;L,wd8,$15;Z;M,$8,$1"
            s = s & "6;L,$8,$13;M,$9,$13;L,$9,$16;M,$4,$14;L,$4,$18;M,$5,$18;L,$5,$14|$4,$13,$5,b"
        Case 99      ' ellipseRibbon2
            s = "25000,50000,12500|p,0,@1,100000;p,25000,@2,75000;a,100000,0,$0;m,$2,1,2;a,$0,0,$3;x,0,$4;p,$5,@3,$0;m,w,$1,200000;a,hc,0,$7;a,$8,wd8,0;a,r,0,$9;a,r,0,$8;a,r,0,wd8;m,h,$6,100000;m,4,$13,w;m,$9,$9,w;a,$9,0,$15;m,$14,$16,1;a,b,0,$17;m,$9,1,2;m,$14,$19,1;a,b,0,$20;a,r,0,$19;m,h,$0,100000;a,$23,0,$13;m,$8,$8,w;a,$8,0,$25;m,$14,$26,1;a,$27,$24,0;a,b,0,$28;a,$13,$24,$28;a,$30,$13,0;a,$31,$24,0;a,b,0,$32;a,b,0,$23;m,$13,14,16;d,$35,$34,2;a,b,0,$36;a,$27,$34,0;a,b,0,$38;a,$28,$34,0;a,b,0,$40;m,$8,1,2;m,$14,$42,1;a,$43,$34,0;a,b,0,$44;a,r,0,$42;a,$32,$34,0;a,b,0,$47;a,$17,$24,0;a,b,0,$49;a,$23,$23,$49;a,b,0,$51|f,0,0,0:M,l,b;Q,$19,$21,$9,$18;L,$8,$29;Q,hc,$33,$11,$29;L,$10,$18;Q,$22,$21,r,b;L,$12,$37;L,r,$23;Q,$46,$45,$11,$39;L,$11,$41;Q,hc,$48,$8,$41;L,$8,$39;Q,$42,$45,l,$23;L,wd8,$37;Z~f,0,0,0:M,$9,$50;L,$9,$18;L,$8,$29;Q,hc,$33,$11,$29;L,$10,$18;L,$10,$50;Q,hc,$52,$9,$50;Z~n,1,0,0:M,l,b;L,wd8,$"
            s = s & "37;L,l,$23;Q,$42,$45,$8,$39;L,$8,$41;Q,hc,$48,$11,$41;L,$11,$39;Q,$46,$45,r,$23;L,$12,$37;L,r,b;Q,$22,$21,$10,$18;L,$11,$29;Q,hc,$33,$8,$29;L,$9,$18;Q,$19,$21,l,b;Z;M,$8,$29;L,$8,$39;M,$11,$39;L,$11,$29;M,$9,$50;L,$9,$18;M,$10,$18;L,$10,$50|$8,$41,$11,$34"
        Case 100      ' ellipseRibbon
            s = "25000,50000,12500|p,0,@1,100000;p,25000,@2,75000;a,100000,0,$0;m,$2,1,2;a,$0,0,$3;x,0,$4;p,$5,@3,$0;m,w,$1,200000;a,hc,0,$7;a,$8,wd8,0;a,r,0,$9;a,r,0,$8;a,r,0,wd8;m,h,$6,100000;m,4,$13,w;m,$9,$9,w;a,$9,0,$15;m,$14,$16,1;m,$9,1,2;m,$14,$18,1;a,r,0,$18;m,h,$0,100000;a,$21,0,$13;m,$8,$8,w;a,$8,0,$23;m,$14,$24,1;a,$25,$22,0;a,$13,$22,$26;a,$27,$13,0;a,$28,$22,0;a,b,0,$21;m,$13,14,16;d,$31,$30,2;a,$25,$30,0;a,$26,$30,0;m,$8,1,2;m,$14,$35,1;a,$36,$30,0;a,r,0,$35;a,$29,$30,0;a,$17,$22,0;a,$21,$21,$40;a,b,0,$13|f,0,0,0:M,l,t;Q,$18,$19,$9,$17;L,$8,$26;Q,hc,$29,$11,$26;L,$10,$17;Q,$20,$19,r,t;L,$12,$32;L,r,$30;Q,$38,$37,$11,$33;L,$11,$34;Q,hc,$39,$8,$34;L,$8,$33;Q,$35,$37,l,$30;L,wd8,$32;Z~f,0,0,0:M,$9,$40;L,$9,$17;L,$8,$26;Q,hc,$29,$11,$26;L,$10,$17;L,$10,$40;Q,hc,$41,$9,$40;Z~n,1,0,0:M,l,t;Q,$18,$19,$9,$17;L,$8,$26;Q,hc,$29,$11,$26;L,$10,$17;Q,$20,$19,r,t;L,$12,$32;L,r,$30;Q,$38,$37,$11,$33;L,$1"
            s = s & "1,$34;Q,hc,$39,$8,$34;L,$8,$33;Q,$35,$37,l,$30;L,wd8,$32;Z;M,$8,$33;L,$8,$26;M,$11,$26;L,$11,$33;M,$9,$17;L,$9,$40;M,$10,$40;L,$10,$17|$8,$21,$11,$34"
        Case 101      ' verticalScroll
            s = "12500|p,0,@1,25000;m,ss,$0,100000;m,$1,1,2;m,$1,1,4;a,$1,$2,0;a,$1,$1,0;a,r,0,$1;a,r,0,$2;a,$6,0,$2;a,b,0,$1;a,b,0,$2|f,0,0,0:M,$2,b;A,$2,$2,cd4,-5400000;L,$2,$10;A,$3,$3,cd4,-10800000;L,$1,$9;L,$1,$2;A,$2,$2,cd2,cd4;L,$7,t;A,$2,$2,k34,cd2;L,$6,$1;L,$6,$10;A,$2,$2,0,cd4;Z;M,$5,$2;A,$2,$2,0,cd4;A,$3,$3,cd4,cd2;Z~f,0,0,0:M,$5,$2;A,$2,$2,0,cd4;A,$3,$3,cd4,cd2;Z;M,$1,$10;A,$2,$2,0,k34;A,$3,$3,k34,cd2;Z~n,1,0,0:M,$1,$9;L,$1,$2;A,$2,$2,cd2,cd4;L,$7,t;A,$2,$2,k34,cd2;L,$6,$1;L,$6,$10;A,$2,$2,0,cd4;L,$2,b;A,$2,$2,cd4,cd2;Z;M,$4,t;A,$2,$2,k34,cd2;A,$3,$3,cd4,cd2;L,$5,$2;M,$6,$1;L,$4,$1;M,$2,$9;A,$3,$3,k34,cd2;L,$1,$10;M,$2,b;A,$2,$2,cd4,-5400000;L,$1,$9|$1,$1,$6,$10"
        Case 102      ' horizontalScroll
            s = "12500|p,0,@1,25000;m,ss,$0,100000;m,$1,1,2;m,$1,1,4;a,$1,$2,0;a,$1,$1,0;a,b,0,$1;a,b,0,$2;a,$6,0,$2;a,r,0,$1;a,r,0,$2|f,0,0,0:M,r,$2;A,$2,$2,0,cd4;L,$10,$2;A,$3,$3,0,cd2;L,$9,$1;L,$2,$1;A,$2,$2,k34,-5400000;L,l,$7;A,$2,$2,cd2,-10800000;L,$1,$6;L,$10,$6;A,$2,$2,cd4,-5400000;Z;M,$2,$5;A,$2,$2,cd4,-5400000;A,$3,$3,0,-10800000;Z~f,0,0,0:M,$2,$5;A,$2,$2,cd4,-5400000;A,$3,$3,0,-10800000;Z;M,$10,$1;A,$2,$2,cd4,-16200000;A,$3,$3,cd2,-10800000;Z~n,1,0,0:M,l,$4;A,$2,$2,cd2,cd4;L,$9,$1;L,$9,$2;A,$2,$2,cd2,cd2;L,r,$8;A,$2,$2,0,cd4;L,$1,$6;L,$1,$7;A,$2,$2,0,cd2;Z;M,$9,$1;L,$10,$1;A,$2,$2,cd4,-5400000;M,$10,$1;L,$10,$2;A,$3,$3,0,cd2;M,$2,$5;L,$2,$4;A,$3,$3,cd2,cd2;A,$2,$2,0,cd2;M,$1,$4;L,$1,$6|$1,$1,$10,$6"
        Case 103      ' wave
            s = "12500,0|p,0,@1,20000;p,-10000,@2,10000;m,h,$0,100000;m,$2,10,3;a,$2,0,$3;a,$2,$3,0;a,b,0,$2;a,$6,0,$3;a,$6,$3,0;m,w,$1,100000;m,w,$1,50000;b,$9;q,$10,0,$10;a,l,0,$12;q,$10,$10,0;a,r,0,$14;d,$12,$15,3;a,$13,$16,0;d,$17,$15,2;a,l,$14,0;a,r,$12,0;a,$19,$16,0;d,$21,$20,2;a,r,0,$11;a,hc,$9,0;a,hc,0,$9;x,$13,$19;n,$15,$20;m,h,$0,50000;a,b,0,$28|f,1,0,0:M,$13,$2;C,$17,$4,$18,$5,$15,$2;L,$20,$6;C,$22,$8,$21,$7,$19,$6;Z|$26,$28,$27,$29"
        Case 104      ' doubleWave
            s = "6250,0|p,0,@1,12500;p,-10000,@2,10000;m,h,$0,100000;m,$2,10,3;a,$2,0,$3;a,$2,$3,0;a,b,0,$2;a,$6,0,$3;a,$6,$3,0;m,w,$1,100000;m,w,$1,50000;b,$9;q,$10,0,$10;a,l,0,$12;q,$10,$10,0;a,r,0,$14;d,$12,$15,6;a,$13,$16,0;d,$12,$15,3;a,$13,$18,0;d,$13,$15,2;a,$20,$16,0;d,$21,$15,2;a,l,$14,0;a,r,$12,0;a,$23,$16,0;a,$23,$18,0;d,$23,$24,2;a,$27,$16,0;d,$28,$24,2;a,r,0,$11;a,hc,$9,0;x,$13,$23;n,$15,$24;m,h,$0,50000;a,b,0,$34|f,1,0,0:M,$13,$2;C,$17,$4,$19,$5,$20,$2;C,$21,$4,$22,$5,$15,$2;L,$24,$6;C,$29,$8,$28,$7,$27,$6;C,$26,$8,$25,$7,$23,$6;Z|$32,$34,$33,$35"
        Case 105      ' wedgeRectCallout
            s = "-20833,62500|m,w,@1,100000;m,h,@2,100000;a,hc,$0,0;a,vc,$1,0;a,$2,0,hc;a,$3,0,vc;m,$0,h,w;b,$1;b,$6;a,$7,0,$8;q,$0,7,2;q,$0,10,5;m,w,$10,12;m,w,$11,12;q,$1,7,2;q,$1,10,5;m,h,$14,12;m,h,$15,12;q,$0,l,$2;q,$9,l,$18;q,$1,$12,$2;q,$9,$20,$12;q,$0,$2,r;q,$9,r,$22;q,$1,$2,$12;q,$9,$24,$12;q,$0,$16,$3;q,$9,$16,$26;q,$1,t,$3;q,$9,$28,t;q,$0,$3,$16;q,$9,$16,$30;q,$1,$3,b;q,$9,$32,b|f,1,0,0:M,l,t;L,$12,t;L,$21,$29;L,$13,t;L,r,t;L,r,$16;L,$23,$31;L,r,$17;L,r,b;L,$13,b;L,$25,$33;L,$12,b;L,l,b;L,l,$17;L,$19,$27;L,l,$16;Z|l,t,r,b"
        Case 106, 137      ' wedgeRoundRectCallout
            s = "-20833,62500,16667|m,w,@1,100000;m,h,@2,100000;a,hc,$0,0;a,vc,$1,0;m,$0,h,w;b,$1;b,$4;a,$5,0,$6;q,$0,7,2;q,$0,10,5;m,w,$8,12;m,w,$9,12;q,$1,7,2;q,$1,10,5;m,h,$12,12;m,h,$13,12;q,$0,l,$2;q,$7,l,$16;q,$1,$10,$2;q,$7,$18,$10;q,$0,$2,r;q,$7,r,$20;q,$1,$2,$10;q,$7,$22,$10;q,$0,$14,$3;q,$7,$14,$24;q,$1,t,$3;q,$7,$26,t;q,$0,$3,$14;q,$7,$14,$28;q,$1,$3,b;q,$7,$30,b;m,ss,@3,100000;a,r,0,$32;a,b,0,$32;m,$32,29289,100000;a,r,0,$35;a,b,0,$35|f,1,0,0:M,l,$32;A,$32,$32,cd2,cd4;L,$10,t;L,$19,$27;L,$11,t;L,$33,t;A,$32,$32,k34,cd4;L,r,$14;L,$21,$29;L,r,$15;L,r,$34;A,$32,$32,0,cd4;L,$11,b;L,$23,$31;L,$10,b;L,$32,b;A,$32,$32,cd4,cd4;L,l,$15;L,$17,$25;L,l,$14;Z|$35,$35,$36,$37"
        Case 107      ' wedgeEllipseCallout
            s = "-20833,62500|m,w,@1,100000;m,h,@2,100000;a,hc,$0,0;a,vc,$1,0;m,$0,h,1;m,$1,w,1;t,$4,$5;a,$6,660000,0;a,$6,0,660000;o,wd2,$7;i,hd2,$7;a,hc,$9,0;a,vc,$10,0;o,wd2,$8;i,hd2,$8;a,hc,$13,0;a,vc,$14,0;t,$9,$10;t,$13,$14;a,$18,0,$17;a,$19,21600000,0;q,$19,$19,$20;o,wd2,2700000;i,hd2,2700000;a,hc,0,$22;a,hc,$22,0;a,vc,0,$23;a,vc,$23,0|f,1,0,0:M,$2,$3;L,$11,$12;A,wd2,hd2,$17,$21;Z|$24,$26,$25,$27"
        Case 108      ' cloudCallout
            s = "-20833,62500|m,w,@1,100000;m,h,@2,100000;a,hc,$0,0;a,vc,$1,0;c,hd2,$0,$1;s,wd2,$0,$1;c,wd2,$4,$5;s,hd2,$4,$5;a,hc,$6,0;a,vc,$7,0;a,$8,0,$2;a,$9,0,$3;h,$10,$11,0;m,ss,6600,21600;a,$12,0,$13;m,$14,1,3;m,ss,1800,21600;a,$15,$16,0;m,$17,$10,$12;m,$17,$11,$12;a,$18,$2,0;a,$19,$3,0;m,ss,4800,21600;m,$15,2,1;a,$22,$23,0;m,$24,$10,$12;m,$24,$11,$12;a,$25,$2,0;a,$26,$3,0;m,ss,1200,21600;m,ss,600,21600;a,$2,$30,0;a,$20,$29,0;a,$27,$16,0;m,w,2977,21600;m,h,3262,21600;m,w,17087,21600;m,h,17337,21600;m,w,67,21600;m,h,21577,21600;m,w,21582,21600;m,h,1235,21600;t,$0,$1|f,1,43200,43200:M,3900,14370;A,6753,9190,-11429249,7426832;A,5333,7267,-8646143,5396714;A,4365,5945,-8748475,5983381;A,4857,6595,-7859164,7034504;A,5333,7273,-4722533,6541615;A,6775,9220,-2776035,7816140;A,5785,7867,37501,6842000;A,6752,9215,1347096,6910353;A,7720,10543,3974558,4542661;A,4360,5918,-16496525,8804134;A,4345,5945,-14809710,"
            s = s & "9151131;Z~f,1,0,0:M,$31,$3;A,$30,$30,0,21600000;Z~f,1,0,0:M,$32,$21;A,$29,$29,0,21600000;Z~f,1,0,0:M,$33,$28;A,$16,$16,0,21600000;Z~n,1,43200,43200:M,4693,26177;A,4345,5945,5204520,1585770;M,6928,34899;A,4360,5918,4416628,686848;M,16478,39090;A,6752,9215,8257449,844866;M,28827,34751;A,6752,9215,387196,959901;M,34129,22954;A,5785,7867,-4217541,4255042;M,41798,15354;A,5333,7273,1819082,1665090;M,38324,5426;A,4857,6595,-824660,891534;M,29078,3952;A,4857,6595,-8950887,1091722;M,22141,4720;A,4365,5945,-9809656,1061181;M,14000,5192;A,6753,9190,-4002417,739161;M,4127,15789;A,6753,9190,9459261,711490|$34,$35,$36,$37"
        Case 109      ' borderCallout1
            s = "18750,-8333,112500,-38333|m,h,@1,100000;m,w,@2,100000;m,h,@3,100000;m,w,@4,100000|f,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z~n,1,0,0:M,$1,$0;L,$3,$2|l,t,r,b"
        Case 110      ' borderCallout2
            s = "18750,-8333,18750,-16667,112500,-46667|m,h,@1,100000;m,w,@2,100000;m,h,@3,100000;m,w,@4,100000;m,h,@5,100000;m,w,@6,100000|f,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z~n,1,0,0:M,$1,$0;L,$3,$2;L,$5,$4|l,t,r,b"
        Case 111, 112      ' borderCallout3
            s = "18750,-8333,18750,-16667,100000,-16667,112963,-8333|m,h,@1,100000;m,w,@2,100000;m,h,@3,100000;m,w,@4,100000;m,h,@5,100000;m,w,@6,100000;m,h,@7,100000;m,w,@8,100000|f,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z~n,1,0,0:M,$1,$0;L,$3,$2;L,$5,$4;L,$7,$6|l,t,r,b"
        Case 113      ' accentCallout1
            s = "18750,-8333,112500,-38333|m,h,@1,100000;m,w,@2,100000;m,h,@3,100000;m,w,@4,100000|f,0,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z~n,1,0,0:M,$1,t;Z;L,$1,b~n,1,0,0:M,$1,$0;L,$3,$2|l,t,r,b"
        Case 114      ' accentCallout2
            s = "18750,-8333,18750,-16667,112500,-46667|m,h,@1,100000;m,w,@2,100000;m,h,@3,100000;m,w,@4,100000;m,h,@5,100000;m,w,@6,100000|f,0,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z~n,1,0,0:M,$1,t;Z;L,$1,b~n,1,0,0:M,$1,$0;L,$3,$2;L,$5,$4|l,t,r,b"
        Case 115, 116      ' accentCallout3
            s = "18750,-8333,18750,-16667,100000,-16667,112963,-8333|m,h,@1,100000;m,w,@2,100000;m,h,@3,100000;m,w,@4,100000;m,h,@5,100000;m,w,@6,100000;m,h,@7,100000;m,w,@8,100000|f,0,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z~n,1,0,0:M,$1,t;Z;L,$1,b~n,1,0,0:M,$1,$0;L,$3,$2;L,$5,$4;L,$7,$6|l,t,r,b"
        Case 117      ' callout1
            s = "18750,-8333,112500,-38333|m,h,@1,100000;m,w,@2,100000;m,h,@3,100000;m,w,@4,100000|f,0,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z~n,1,0,0:M,$1,$0;L,$3,$2|l,t,r,b"
        Case 118      ' callout2
            s = "18750,-8333,18750,-16667,112500,-46667|m,h,@1,100000;m,w,@2,100000;m,h,@3,100000;m,w,@4,100000;m,h,@5,100000;m,w,@6,100000|f,0,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z~n,1,0,0:M,$1,$0;L,$3,$2;L,$5,$4|l,t,r,b"
        Case 119, 120      ' callout3
            s = "18750,-8333,18750,-16667,100000,-16667,112963,-8333|m,h,@1,100000;m,w,@2,100000;m,h,@3,100000;m,w,@4,100000;m,h,@5,100000;m,w,@6,100000;m,h,@7,100000;m,w,@8,100000|f,0,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z~n,1,0,0:M,$1,$0;L,$3,$2;L,$5,$4;L,$7,$6|l,t,r,b"
        Case 121      ' accentBorderCallout1
            s = "18750,-8333,112500,-38333|m,h,@1,100000;m,w,@2,100000;m,h,@3,100000;m,w,@4,100000|f,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z~n,1,0,0:M,$1,t;Z;L,$1,b~n,1,0,0:M,$1,$0;L,$3,$2|l,t,r,b"
        Case 122      ' accentBorderCallout2
            s = "18750,-8333,18750,-16667,112500,-46667|m,h,@1,100000;m,w,@2,100000;m,h,@3,100000;m,w,@4,100000;m,h,@5,100000;m,w,@6,100000|f,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z~n,1,0,0:M,$1,t;Z;L,$1,b~n,1,0,0:M,$1,$0;L,$3,$2;L,$5,$4|l,t,r,b"
        Case 123, 124      ' accentBorderCallout3
            s = "18750,-8333,18750,-16667,100000,-16667,112963,-8333|m,h,@1,100000;m,w,@2,100000;m,h,@3,100000;m,w,@4,100000;m,h,@5,100000;m,w,@6,100000;m,h,@7,100000;m,w,@8,100000|f,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z~n,1,0,0:M,$1,t;Z;L,$1,b~n,1,0,0:M,$1,$0;L,$3,$2;L,$5,$4;L,$7,$6|l,t,r,b"
        Case 125      ' actionButtonBlank
            s = "||f,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z|l,t,r,b"
        Case 126      ' actionButtonHome
            s = "|m,ss,3,8;a,vc,0,$0;a,vc,$0,0;a,hc,0,$0;a,hc,$0,0;m,ss,3,4;m,$5,1,16;m,$5,1,8;m,$5,3,16;m,$5,5,16;m,$5,7,16;m,$5,9,16;m,$5,11,16;m,$5,3,4;m,$5,13,16;m,$5,7,8;a,$1,$6,0;a,$1,$8,0;a,$1,$9,0;a,$1,$13,0;a,$3,$7,0;a,$3,$10,0;a,$3,$11,0;a,$3,$12,0;a,$3,$14,0;a,$3,$15,0|f,0,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z;M,hc,$1;L,$3,vc;L,$20,vc;L,$20,$2;L,$25,$2;L,$25,vc;L,$4,vc;L,$24,$18;L,$24,$16;L,$23,$16;L,$23,$17;Z~f,0,0,0:M,$24,$18;L,$24,$16;L,$23,$16;L,$23,$17;Z;M,$20,vc;L,$20,$2;L,$21,$2;L,$21,$19;L,$22,$19;L,$22,$2;L,$25,$2;L,$25,vc;Z~f,0,0,0:M,hc,$1;L,$3,vc;L,$4,vc;Z;M,$21,$19;L,$22,$19;L,$22,$2;L,$21,$2;Z~n,1,0,0:M,hc,$1;L,$23,$17;L,$23,$16;L,$24,$16;L,$24,$18;L,$4,vc;L,$25,vc;L,$25,$2;L,$20,$2;L,$20,vc;L,$3,vc;Z;M,$23,$17;L,$24,$18;M,$25,vc;L,$20,vc;M,$21,$2;L,$21,$19;L,$22,$19;L,$22,$2~n,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z|l,t,r,b"
        Case 127      ' actionButtonHelp
            s = "|m,ss,3,8;a,vc,0,$0;a,hc,0,$0;m,ss,3,4;m,$3,1,7;m,$3,3,14;m,$3,2,7;m,$3,3,7;m,$3,4,7;m,$3,17,28;m,$3,21,28;m,$3,11,14;a,$1,$6,0;a,$1,$9,0;a,$1,$10,0;a,$1,$11,0;a,$2,$5,0;a,$2,$7,0;a,$2,$8,0;m,$3,1,14;m,$3,3,28|f,0,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z;M,$16,$12;A,$6,$6,cd2,cd2;A,$4,$5,0,cd4;A,$19,$20,k34,-5400000;L,$18,$14;L,$17,$14;L,$17,$13;A,$4,$5,cd2,cd4;A,$19,$20,cd4,-5400000;A,$4,$4,0,-10800000;Z;M,hc,$15;A,$20,$20,k34,21600000;Z~f,0,0,0:M,$16,$12;A,$6,$6,cd2,cd2;A,$4,$5,0,cd4;A,$19,$20,k34,-5400000;L,$18,$14;L,$17,$14;L,$17,$13;A,$4,$5,cd2,cd4;A,$19,$20,cd4,-5400000;A,$4,$4,0,-10800000;Z;M,hc,$15;A,$20,$20,k34,21600000;Z~n,1,0,0:M,$16,$12;A,$6,$6,cd2,cd2;A,$4,$5,0,cd4;A,$19,$20,k34,-5400000;L,$18,$14;L,$17,$14;L,$17,$13;A,$4,$5,cd2,cd4;A,$19,$20,cd4,-5400000;A,$4,$4,0,-10800000;Z;M,hc,$15;A,$20,$20,k34,21600000;Z~n,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z|l,t,r,b"
        Case 128      ' actionButtonInformation
            s = "|m,ss,3,8;a,vc,0,$0;a,hc,0,$0;m,ss,3,4;m,$3,1,32;m,$3,5,16;m,$3,3,8;m,$3,13,32;m,$3,19,32;m,$3,11,16;m,$3,13,16;m,$3,7,8;a,$1,$4,0;a,$1,$5,0;a,$1,$6,0;a,$1,$10,0;a,$1,$11,0;a,$2,$5,0;a,$2,$7,0;a,$2,$8,0;a,$2,$9,0;m,$3,3,32|f,0,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z;M,hc,$1;A,$0,$0,k34,21600000;Z~f,0,0,0:M,hc,$1;A,$0,$0,k34,21600000;Z;M,hc,$12;A,$21,$21,k34,21600000;M,$17,$13;L,$17,$14;L,$18,$14;L,$18,$15;L,$17,$15;L,$17,$16;L,$20,$16;L,$20,$15;L,$19,$15;L,$19,$13;Z~f,0,0,0:M,hc,$12;A,$21,$21,k34,21600000;M,$17,$13;L,$19,$13;L,$19,$15;L,$20,$15;L,$20,$16;L,$17,$16;L,$17,$15;L,$18,$15;L,$18,$14;L,$17,$14;Z~n,1,0,0:M,hc,$1;A,$0,$0,k34,21600000;Z;M,hc,$12;A,$21,$21,k34,21600000;M,$17,$13;L,$19,$13;L,$19,$15;L,$20,$15;L,$20,$16;L,$17,$16;L,$17,$15;L,$18,$15;L,$18,$14;L,$17,$14;Z~n,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z|l,t,r,b"
        Case 129      ' actionButtonBackPrevious
            s = "|m,ss,3,8;a,vc,0,$0;a,vc,$0,0;a,hc,0,$0;a,hc,$0,0|f,0,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z;M,$3,vc;L,$4,$1;L,$4,$2;Z~f,0,0,0:M,$3,vc;L,$4,$1;L,$4,$2;Z~n,1,0,0:M,$3,vc;L,$4,$1;L,$4,$2;Z~n,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z|l,t,r,b"
        Case 130      ' actionButtonForwardNext
            s = "|m,ss,3,8;a,vc,0,$0;a,vc,$0,0;a,hc,0,$0;a,hc,$0,0|f,0,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z;M,$4,vc;L,$3,$1;L,$3,$2;Z~f,0,0,0:M,$4,vc;L,$3,$1;L,$3,$2;Z~n,1,0,0:M,$4,vc;L,$3,$2;L,$3,$1;Z~n,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z|l,t,r,b"
        Case 131      ' actionButtonBeginning
            s = "|m,ss,3,8;a,vc,0,$0;a,vc,$0,0;a,hc,0,$0;a,hc,$0,0;m,ss,3,4;m,$5,1,8;m,$5,1,4;a,$3,$6,0;a,$3,$7,0|f,0,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z;M,$9,vc;L,$4,$1;L,$4,$2;Z;M,$8,$1;L,$3,$1;L,$3,$2;L,$8,$2;Z~f,0,0,0:M,$9,vc;L,$4,$1;L,$4,$2;Z;M,$8,$1;L,$3,$1;L,$3,$2;L,$8,$2;Z~n,1,0,0:M,$9,vc;L,$4,$1;L,$4,$2;Z;M,$8,$1;L,$8,$2;L,$3,$2;L,$3,$1;Z~n,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z|l,t,r,b"
        Case 132      ' actionButtonEnd
            s = "|m,ss,3,8;a,vc,0,$0;a,vc,$0,0;a,hc,0,$0;a,hc,$0,0;m,ss,3,4;m,$5,3,4;m,$5,7,8;a,$3,$6,0;a,$3,$7,0|f,0,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z;M,$8,vc;L,$3,$1;L,$3,$2;Z;M,$9,$1;L,$4,$1;L,$4,$2;L,$9,$2;Z~f,0,0,0:M,$8,vc;L,$3,$1;L,$3,$2;Z;M,$9,$1;L,$4,$1;L,$4,$2;L,$9,$2;Z~n,1,0,0:M,$8,vc;L,$3,$2;L,$3,$1;Z;M,$9,$1;L,$4,$1;L,$4,$2;L,$9,$2;Z~n,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z|l,t,r,b"
        Case 133      ' actionButtonReturn
            s = "|m,ss,3,8;a,vc,0,$0;a,vc,$0,0;a,hc,0,$0;a,hc,$0,0;m,ss,3,4;m,$5,7,8;m,$5,3,4;m,$5,5,8;m,$5,3,8;m,$5,1,4;a,$1,$7,0;a,$1,$8,0;a,$1,$10,0;a,$3,$6,0;a,$3,$7,0;a,$3,$8,0;a,$3,$9,0;a,$3,$10,0;m,$5,1,8|f,0,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z;M,$4,$13;L,$15,$1;L,hc,$13;L,$16,$13;L,$16,$12;A,$19,$19,0,cd4;L,$17,$11;A,$19,$19,cd4,cd4;L,$18,$13;L,$3,$13;L,$3,$12;A,$9,$9,cd2,-5400000;L,hc,$2;A,$9,$9,cd4,-5400000;L,$14,$13;Z~f,0,0,0:M,$4,$13;L,$15,$1;L,hc,$13;L,$16,$13;L,$16,$12;A,$19,$19,0,cd4;L,$17,$11;A,$19,$19,cd4,cd4;L,$18,$13;L,$3,$13;L,$3,$12;A,$9,$9,cd2,-5400000;L,hc,$2;A,$9,$9,cd4,-5400000;L,$14,$13;Z~n,1,0,0:M,$4,$13;L,$14,$13;L,$14,$12;A,$9,$9,0,cd4;L,$17,$2;A,$9,$9,cd4,cd4;L,$3,$13;L,$18,$13;L,$18,$12;A,$19,$19,cd2,-5400000;L,hc,$11;A,$19,$19,cd4,-5400000;L,$16,$13;L,hc,$13;L,$15,$1;Z~n,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z|l,t,r,b"
        Case 134      ' actionButtonDocument
            s = "|m,ss,3,8;a,vc,0,$0;a,vc,$0,0;m,ss,9,32;a,hc,0,$3;a,hc,$3,0;m,ss,3,16;a,$5,0,$6;a,$1,$6,0|f,0,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z;M,$4,$1;L,$7,$1;L,$5,$8;L,$5,$2;L,$4,$2;Z~f,0,0,0:M,$4,$1;L,$7,$1;L,$7,$8;L,$5,$8;L,$5,$2;L,$4,$2;Z~f,0,0,0:M,$7,$1;L,$7,$8;L,$5,$8;Z~n,1,0,0:M,$4,$1;L,$7,$1;L,$5,$8;L,$5,$2;L,$4,$2;Z;M,$5,$8;L,$7,$8;L,$7,$1~n,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z|l,t,r,b"
        Case 135      ' actionButtonSound
            s = "|m,ss,3,8;a,vc,0,$0;a,vc,$0,0;a,hc,0,$0;a,hc,$0,0;m,ss,3,4;m,$5,1,8;m,$5,5,16;m,$5,5,8;m,$5,11,16;m,$5,3,4;m,$5,7,8;a,$1,$6,0;a,$1,$7,0;a,$1,$9,0;a,$1,$11,0;a,$3,$7,0;a,$3,$8,0;a,$3,$10,0|f,0,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z;M,$3,$13;L,$3,$14;L,$16,$14;L,$17,$2;L,$17,$1;L,$16,$13;Z~f,0,0,0:M,$3,$13;L,$3,$14;L,$16,$14;L,$17,$2;L,$17,$1;L,$16,$13;Z~n,1,0,0:M,$3,$13;L,$16,$13;L,$17,$1;L,$17,$2;L,$16,$14;L,$3,$14;Z;M,$18,$13;L,$4,$12;M,$18,vc;L,$4,vc;M,$18,$14;L,$4,$15~n,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z|l,t,r,b"
    End Select
    PresetSrc2 = s
End Function

Private Function PresetSrc3(ByVal ast As Long) As String
    Dim s As String
    Select Case ast
        Case 136      ' actionButtonMovie
            s = "|m,ss,3,8;a,vc,0,$0;a,vc,$0,0;a,hc,0,$0;a,hc,$0,0;m,ss,3,4;m,$5,1455,21600;m,$5,1905,21600;m,$5,2325,21600;m,$5,16155,21600;m,$5,17010,21600;m,$5,19335,21600;m,$5,19725,21600;m,$5,20595,21600;m,$5,5280,21600;m,$5,5730,21600;m,$5,6630,21600;m,$5,7492,21600;m,$5,9067,21600;m,$5,9555,21600;m,$5,13342,21600;m,$5,14580,21600;m,$5,15592,21600;a,$3,$6,0;a,$3,$7,0;a,$3,$8,0;a,$3,$9,0;a,$3,$10,0;a,$3,$11,0;a,$3,$12,0;a,$3,$13,0;a,$1,$14,0;a,$1,$15,0;a,$1,$16,0;a,$1,$17,0;a,$1,$18,0;a,$1,$19,0;a,$1,$20,0;a,$1,$21,0;a,$1,$22,0;a,$1,$23,0|f,0,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z;M,$3,$31;L,$3,$36;L,$23,$36;L,$24,$35;L,$25,$35;L,$25,$39;L,$27,$39;L,$27,$37;L,$28,$37;L,$30,$38;L,$4,$38;L,$4,$33;L,$30,$33;L,$29,$34;L,$27,$34;L,$27,$33;L,$26,$32;L,$24,$32;L,$23,$31;Z~f,0,0,0:M,$3,$31;L,$3,$36;L,$23,$36;L,$24,$35;L,$25,$35;L,$25,$39;L,$27,$39;L,$27,$37;L,$28,$37;L,$30,$38;L,$4,$38;L,$4,$33;L,$30,$33;L,$29,$34;L"
            s = s & ",$27,$34;L,$27,$33;L,$26,$32;L,$24,$32;L,$23,$31;Z~n,1,0,0:M,$3,$31;L,$23,$31;L,$24,$32;L,$26,$32;L,$27,$33;L,$27,$34;L,$29,$34;L,$30,$33;L,$4,$33;L,$4,$38;L,$30,$38;L,$28,$37;L,$27,$37;L,$27,$39;L,$25,$39;L,$25,$35;L,$24,$35;L,$23,$36;L,$3,$36;Z~n,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z|l,t,r,b"
        Case 139      ' flowChartOfflineStorage
            s = "|m,w,3,4|f,0,2,2:M,0,0;L,2,0;L,1,2;Z~n,1,5,5:M,2,4;L,3,4~n,1,2,2:M,0,0;L,2,0;L,1,2;Z|wd4,t,$0,vc"
        Case 140      ' leftRightRibbon
            s = "50000,50000,16667|p,0,@3,33333;a,100000,0,$0;p,0,@1,$1;a,wd2,0,wd32;m,100000,$3,ss;p,0,@2,$4;m,ss,$5,100000;a,r,0,$6;m,h,$2,200000;m,h,$0,-200000;a,vc,$9,$8;a,vc,$8,$9;a,$10,$8,0;a,b,0,$12;m,$12,2,1;a,b,0,$14;a,$14,0,$10;a,b,0,$16;m,$0,ss,400000;a,hc,0,wd32;a,hc,wd32,0;a,$10,$18,0;a,$17,0,$18|f,0,0,0:M,l,$12;L,$6,t;L,$6,$10;L,hc,$10;A,wd32,$18,k34,cd2;A,wd32,$18,k34,-10800000;L,$7,$17;L,$7,$15;L,r,$13;L,$7,b;L,$7,$11;L,hc,$11;A,wd32,$18,cd4,cd4;L,$19,$16;L,$6,$16;L,$6,$14;Z~f,0,0,0:M,$20,$21;A,wd32,$18,0,cd4;A,wd32,$18,k34,-10800000;L,$20,$17;Z~n,1,0,0:M,l,$12;L,$6,t;L,$6,$10;L,hc,$10;A,wd32,$18,k34,cd2;A,wd32,$18,k34,-10800000;L,$7,$17;L,$7,$15;L,r,$13;L,$7,b;L,$7,$11;L,hc,$11;A,wd32,$18,cd4,cd4;L,$19,$16;L,$6,$16;L,$6,$14;Z;M,$20,$21;L,$20,$17;M,$19,$22;L,$19,$16|$6,$10,$7,$11"
        Case 141      ' diagStripe
            s = "50000|p,0,@1,100000;m,w,$0,100000;m,$1,1,2;d,$1,r,2;m,h,$0,100000;m,$4,1,2;d,$4,b,2|f,1,0,0:M,l,$4;L,$1,t;L,r,t;L,l,b;Z|l,t,$3,$6"
        Case 142      ' pie
            s = "!0,!16200000|p,0,@1,21599999;p,0,@2,21599999;a,$1,0,$0;a,$2,21600000,0;q,$2,$2,$3;i,wd2,$0;o,hd2,$0;c,wd2,$6,$5;s,hd2,$6,$5;a,hc,$7,0;a,vc,$8,0;i,wd2,$1;o,hd2,$1;c,wd2,$12,$11;s,hd2,$12,$11;a,hc,$13,0;a,vc,$14,0;o,wd2,2700000;i,hd2,2700000;a,hc,0,$17;a,hc,$17,0;a,vc,0,$18;a,vc,$18,0|f,1,0,0:M,$9,$10;A,wd2,hd2,$0,$4;L,hc,vc;Z|$19,$21,$20,$22"
        Case 143      ' nonIsoscelesTrapezoid
            s = "25000,25000|m,50000,w,ss;p,0,@1,$0;p,0,@2,$0;m,ss,$1,200000;m,ss,$1,100000;m,ss,$2,100000;a,r,0,$5;d,r,$6,2;m,wd3,$1,$0;x,$1,$2;m,hd3,$9,$0;m,wd3,$2,$0;a,r,0,$11|f,1,0,0:M,l,b;L,$4,t;L,$6,t;L,r,b;Z|$8,$10,$12,b"
        Case 144      ' decagon
            s = "105146|m,hd2,@1,100000;o,wd2,2160000;o,wd2,4320000;a,hc,0,$1;a,hc,0,$2;a,hc,$2,0;a,hc,$1,0;i,$0,4320000;i,$0,2160000;a,vc,0,$7;a,vc,0,$8;a,vc,$8,0;a,vc,$7,0|f,1,0,0:M,l,vc;L,$3,$10;L,$4,$9;L,$5,$9;L,$6,$10;L,r,vc;L,$6,$11;L,$5,$12;L,$4,$12;L,$3,$11;Z|$3,$10,$6,$11"
        Case 145      ' heptagon
            s = "102572,105210|m,wd2,@1,100000;m,hd2,@2,100000;m,vc,@2,100000;m,$0,97493,100000;m,$0,78183,100000;m,$0,43388,100000;m,$1,62349,100000;m,$1,22252,100000;m,$1,90097,100000;a,hc,0,$3;a,hc,0,$4;a,hc,0,$5;a,hc,$5,0;a,hc,$4,0;a,hc,$3,0;a,$2,0,$6;a,$2,$7,0;a,$2,$8,0;a,b,0,$15|f,1,0,0:M,$9,$16;L,$10,$15;L,hc,t;L,$13,$15;L,$14,$16;L,$12,$17;L,$11,$17;Z|$10,$15,$13,$18"
        Case 146      ' dodecagon
            s = "|m,w,2894,21600;m,w,7906,21600;m,w,13694,21600;m,w,18706,21600;m,h,2894,21600;m,h,7906,21600;m,h,13694,21600;m,h,18706,21600|f,1,0,0:M,l,$5;L,$0,$4;L,$1,t;L,$2,t;L,$3,$4;L,r,$5;L,r,$6;L,$3,$7;L,$2,b;L,$1,b;L,$0,$7;L,l,$6;Z|$0,$4,$3,$7"
        Case 147      ' star6
            s = "28868,115470|p,0,@1,50000;m,wd2,@2,100000;o,$1,1800000;a,hc,0,$2;a,hc,$2,0;a,vc,hd4,0;m,$1,$0,50000;m,hd2,$0,50000;m,$6,1,2;a,hc,0,$6;a,hc,0,$8;a,hc,$8,0;a,hc,$6,0;i,$7,3600000;a,vc,0,$13;a,vc,$13,0;a,vc,0,$7|f,1,0,0:M,$3,hd4;L,$10,$14;L,hc,t;L,$11,$14;L,$4,hd4;L,$12,vc;L,$4,$5;L,$11,$15;L,hc,b;L,$10,$15;L,$3,$5;L,$9,vc;Z|$9,$14,$12,$15"
        Case 148      ' star7
            s = "34601,102572,105210|p,0,@1,50000;m,wd2,@2,100000;m,hd2,@3,100000;m,vc,@3,100000;m,$1,97493,100000;m,$1,78183,100000;m,$1,43388,100000;m,$2,62349,100000;m,$2,22252,100000;m,$2,90097,100000;a,hc,0,$4;a,hc,0,$5;a,hc,0,$6;a,hc,$6,0;a,hc,$5,0;a,hc,$4,0;a,$3,0,$7;a,$3,$8,0;a,$3,$9,0;m,$1,$0,50000;m,$2,$0,50000;m,$19,97493,100000;m,$19,78183,100000;m,$19,43388,100000;a,hc,0,$21;a,hc,0,$22;a,hc,0,$23;a,hc,$23,0;a,hc,$22,0;a,hc,$21,0;m,$20,90097,100000;m,$20,22252,100000;m,$20,62349,100000;a,$3,0,$30;a,$3,0,$31;a,$3,$32,0;a,$3,$20,0;a,$3,0,$20|f,1,0,0:M,$10,$17;L,$24,$34;L,$11,$16;L,$26,$33;L,hc,t;L,$27,$33;L,$14,$16;L,$29,$34;L,$15,$17;L,$28,$35;L,$13,$18;L,hc,$36;L,$12,$18;L,$25,$35;Z|$25,$33,$28,$35"
        Case 149      ' star10
            s = "42533,105146|p,0,@1,50000;m,wd2,@2,100000;m,$1,95106,100000;m,$1,58779,100000;a,hc,0,$2;a,hc,0,$3;a,hc,$3,0;a,hc,$2,0;m,hd2,80902,100000;m,hd2,30902,100000;a,vc,0,$8;a,vc,0,$9;a,vc,$9,0;a,vc,$8,0;m,$1,$0,50000;m,hd2,$0,50000;m,$14,80902,100000;m,$14,30902,100000;m,$15,95106,100000;m,$15,58779,100000;a,hc,0,$14;a,hc,0,$16;a,hc,0,$17;a,hc,$17,0;a,hc,$16,0;a,hc,$14,0;a,vc,0,$18;a,vc,0,$19;a,vc,$19,0;a,vc,$18,0;a,vc,0,$15|f,1,0,0:M,$4,$11;L,$21,$27;L,$5,$10;L,$22,$26;L,hc,t;L,$23,$26;L,$6,$10;L,$24,$27;L,$7,$11;L,$25,vc;L,$7,$12;L,$24,$28;L,$6,$13;L,$23,$29;L,hc,b;L,$22,$29;L,$5,$13;L,$21,$28;L,$4,$12;L,$20,vc;Z|$21,$27,$24,$28"
        Case 150      ' star12
            s = "37500|p,0,@1,50000;o,wd2,1800000;i,hd2,3600000;a,hc,0,$1;m,w,3,4;a,hc,$1,0;a,vc,0,$2;m,h,3,4;a,vc,$2,0;m,wd2,$0,50000;m,hd2,$0,50000;o,$9,900000;o,$9,2700000;o,$9,4500000;i,$10,4500000;i,$10,2700000;i,$10,900000;a,hc,0,$11;a,hc,0,$12;a,hc,0,$13;a,hc,$13,0;a,hc,$12,0;a,hc,$11,0;a,vc,0,$14;a,vc,0,$15;a,vc,0,$16;a,vc,$16,0;a,vc,$15,0;a,vc,$14,0;a,vc,0,$10|f,1,0,0:M,l,vc;L,$17,$25;L,$3,hd4;L,$18,$24;L,wd4,$6;L,$19,$23;L,hc,t;L,$20,$23;L,$4,$6;L,$21,$24;L,$5,hd4;L,$22,$25;L,r,vc;L,$22,$26;L,$5,$7;L,$21,$27;L,$4,$8;L,$20,$28;L,hc,b;L,$19,$28;L,wd4,$8;L,$18,$27;L,$3,$7;L,$17,$26;Z|$18,$24,$21,$27"
        Case 151      ' round1Rect
            s = "16667|p,0,@1,50000;m,ss,$0,100000;a,r,0,$1;m,$1,29289,100000;a,r,0,$3|f,1,0,0:M,l,t;L,$2,t;A,$1,$1,k34,cd4;L,r,b;L,l,b;Z|l,t,$4,b"
        Case 152      ' round2SameRect
            s = "16667,0|p,0,@1,50000;p,0,@2,50000;m,ss,$0,100000;a,r,0,$2;m,ss,$1,100000;a,r,0,$4;a,b,0,$4;a,$2,0,$4;m,$2,29289,100000;m,$4,29289,100000;q,$7,$8,$9;a,r,0,$10;a,b,0,$9|f,1,0,0:M,$2,t;L,$3,t;A,$2,$2,k34,cd4;L,r,$6;A,$4,$4,0,cd4;L,$4,b;A,$4,$4,cd4,cd4;L,l,$2;A,$2,$2,cd2,cd4;Z|$10,$8,$11,$12"
        Case 153      ' round2DiagRect
            s = "16667,0|p,0,@1,50000;p,0,@2,50000;m,ss,$0,100000;a,b,0,$2;m,ss,$1,100000;a,r,0,$4;a,b,0,$4;m,$2,29289,100000;m,$4,29289,100000;a,$7,0,$8;q,$9,$7,$8;a,r,0,$10;a,b,0,$10|f,1,0,0:M,$2,t;L,$5,t;A,$4,$4,k34,cd4;L,r,$3;A,$2,$2,0,cd4;L,$4,b;A,$4,$4,cd4,cd4;L,l,$2;A,$2,$2,cd2,cd4;Z|$10,$10,$11,$12"
        Case 154      ' snipRoundRect
            s = "16667,16667|p,0,@1,50000;p,0,@2,50000;m,ss,$0,100000;m,ss,$1,100000;a,r,0,$3;m,$2,29289,100000;d,$4,r,2|f,1,0,0:M,$2,t;L,$4,t;L,r,$3;L,r,b;L,l,b;L,l,$2;A,$2,$2,cd2,cd4;Z|$5,$5,$6,b"
        Case 155      ' snip1Rect
            s = "16667|p,0,@1,50000;m,ss,$0,100000;a,r,0,$1;m,$1,1,2;d,$2,r,2|f,1,0,0:M,l,t;L,$2,t;L,r,$1;L,r,b;L,l,b;Z|l,$3,$4,b"
        Case 156      ' snip2SameRect
            s = "16667,0|p,0,@1,50000;p,0,@2,50000;m,ss,$0,100000;a,r,0,$2;m,ss,$1,100000;a,r,0,$4;a,b,0,$4;a,$2,0,$4;q,$7,$2,$4;m,$8,1,2;a,r,0,$9;m,$2,1,2;d,$6,b,2|f,1,0,0:M,$2,t;L,$3,t;L,r,$2;L,r,$6;L,$5,b;L,$4,b;L,l,$6;L,l,$2;Z|$9,$11,$10,$12"
        Case 157      ' snip2DiagRect
            s = "0,16667|p,0,@1,50000;p,0,@2,50000;m,ss,$0,100000;a,r,0,$2;a,b,0,$2;m,ss,$1,100000;a,r,0,$5;a,b,0,$5;a,$2,0,$5;q,$8,$2,$5;m,$9,1,2;a,r,0,$10;a,b,0,$10|f,1,0,0:M,$2,t;L,$6,t;L,r,$5;L,r,$4;L,$3,b;L,$5,b;L,l,$7;L,l,$2;Z|$10,$10,$11,$12"
        Case 158      ' frame
            s = "12500|p,0,@1,50000;m,ss,$0,100000;a,r,0,$1;a,b,0,$1|f,1,0,0:M,l,t;L,r,t;L,r,b;L,l,b;Z;M,$1,$1;L,$1,$3;L,$2,$3;L,$2,$1;Z|$1,$1,$2,$3"
        Case 159      ' halfFrame
            s = "33333,33333|m,100000,w,ss;p,0,@2,$0;m,ss,$1,100000;m,h,$2,w;a,h,0,$3;m,100000,$4,ss;p,0,@1,$5;m,ss,$6,100000;m,$7,w,h;a,r,0,$8;m,$2,h,w;a,b,0,$10;m,$2,1,2;d,$11,b,2;d,$9,r,2;m,$7,1,2|f,1,0,0:M,l,t;L,r,t;L,$9,$7;L,$2,$7;L,$2,$11;L,l,b;Z|l,t,r,b"
        Case 160      ' teardrop
            s = "100000|p,0,@1,200000;r,2;m,wd2,$1,1;m,hd2,$1,1;m,$2,$0,100000;m,$3,$0,100000;o,$4,2700000;i,$5,2700000;a,hc,$6,0;a,vc,0,$7;d,hc,$8,2;d,vc,$9,2;o,wd2,2700000;i,hd2,2700000;a,hc,0,$12;a,hc,$12,0;a,vc,0,$13;a,vc,$13,0|f,1,0,0:M,l,vc;A,wd2,hd2,cd2,cd4;Q,$10,t,$8,$9;Q,r,$11,r,vc;A,wd2,hd2,0,cd4;A,wd2,hd2,cd4,cd4;Z|$14,$16,$15,$17"
        Case 161      ' chord
            s = "!2700000,!16200000|p,0,@1,21599999;p,0,@2,21599999;a,$1,0,$0;a,$2,21600000,0;q,$2,$2,$3;i,wd2,$0;o,hd2,$0;c,wd2,$6,$5;s,hd2,$6,$5;i,wd2,$1;o,hd2,$1;c,wd2,$10,$9;s,hd2,$10,$9;a,hc,$7,0;a,vc,$8,0;a,hc,$11,0;a,vc,$12,0;d,$13,$15,2;d,$14,$16,2;m,$4,1,2;a,$0,$19,cd2;o,wd2,2700000;i,hd2,2700000;a,hc,0,$21;a,hc,$21,0;a,vc,0,$22;a,vc,$22,0|f,1,0,0:M,$13,$14;A,wd2,hd2,$0,$4;Z|$23,$25,$24,$26"
        Case 162      ' corner
            s = "50000,50000|m,100000,h,ss;m,100000,w,ss;p,0,@1,$0;p,0,@2,$1;m,ss,$3,100000;m,ss,$2,100000;a,b,0,$5;m,$4,1,2;d,$6,b,2;a,w,0,h;q,$9,$6,t;q,$9,r,$4|f,1,0,0:M,l,t;L,$4,t;L,$4,$6;L,r,$6;L,r,b;L,l,b;Z|l,$10,$11,b"
        Case 163      ' mathPlus
            s = "23520|p,0,@1,73490;m,w,73490,200000;m,h,73490,200000;m,ss,$0,200000;a,hc,0,$1;a,hc,0,$3;a,hc,$3,0;a,hc,$1,0;a,vc,0,$2;a,vc,0,$3;a,vc,$3,0;a,vc,$2,0|f,1,0,0:M,$4,$9;L,$5,$9;L,$5,$8;L,$6,$8;L,$6,$9;L,$7,$9;L,$7,$10;L,$6,$10;L,$6,$11;L,$5,$11;L,$5,$10;L,$4,$10;Z|$4,$9,$7,$10"
        Case 164      ' mathMinus
            s = "23520|p,0,@1,100000;m,h,$0,200000;m,w,73490,200000;a,vc,0,$1;a,vc,$1,0;a,hc,0,$2;a,hc,$2,0|f,1,0,0:M,$5,$3;L,$6,$3;L,$6,$4;L,$5,$4;Z|$5,$3,$6,$4"
        Case 165      ' mathMultiply
            s = "23520|p,0,@1,51965;m,ss,$0,100000;t,w,h;i,1,$2;o,1,$2;g,1,$2;h,w,h,0;m,$6,51965,100000;a,$6,0,$7;m,$4,$8,2;m,$3,$8,2;m,$3,$1,2;m,$4,$1,2;a,$9,0,$11;a,$10,$12,0;a,$9,$11,0;a,$10,0,$12;a,hc,0,$15;m,$17,$5,1;a,$18,$16,0;a,r,0,$15;a,r,0,$13;a,vc,0,$14;m,$22,1,$5;a,$21,0,$23;a,$13,$23,0;a,b,0,$14;a,b,0,$16;a,b,0,$19;a,r,0,$9;a,b,0,$10|f,1,0,0:M,$13,$14;L,$15,$16;L,hc,$19;L,$20,$16;L,$21,$14;L,$24,vc;L,$21,$26;L,$20,$27;L,hc,$28;L,$15,$27;L,$13,$26;L,$25,vc;Z|$13,$16,$21,$27"
        Case 166      ' mathDivide
            s = "23520,5880,11760|p,1000,@1,36745;a,0,0,$0;d,73490,$1,4;m,36745,w,h;n,$2,$3;p,1000,@3,$4;m,-4,$5,1;a,73490,$6,$0;p,0,@2,$7;m,h,$0,200000;m,h,$8,100000;m,h,$5,100000;m,w,73490,200000;a,vc,0,$9;a,vc,$9,0;a,$10,$11,0;a,$13,0,$15;a,$16,0,$11;a,b,0,$17;a,hc,0,$12;a,hc,$12,0;a,hc,0,$11|f,1,0,0:M,hc,$17;A,$11,$11,k34,21600000;Z;M,hc,$18;A,$11,$11,cd4,21600000;Z;M,$19,$13;L,$20,$13;L,$20,$14;L,$19,$14;Z|$19,$13,$20,$14"
        Case 167      ' mathEqual
            s = "23520,11760|p,0,@1,36745;m,$0,2,1;a,100000,0,$1;p,0,@2,$2;m,h,$0,100000;m,h,$3,200000;m,w,73490,200000;a,vc,0,$5;a,vc,$5,0;a,$7,0,$4;a,$8,$4,0;a,hc,0,$6;a,hc,$6,0;d,$9,$7,2;d,$8,$10,2|f,1,0,0:M,$11,$9;L,$12,$9;L,$12,$7;L,$11,$7;Z;M,$11,$8;L,$12,$8;L,$12,$10;L,$11,$10;Z|$11,$9,$12,$10"
        Case 168      ' mathNotEqual
            s = "23520,!6600000,11760|p,0,@1,50000;p,4200000,@2,6600000;m,$0,2,1;a,100000,0,$2;p,0,@3,$3;m,h,$0,100000;m,h,$4,200000;m,w,73490,200000;a,hc,0,$7;a,hc,$7,0;a,vc,0,$6;a,vc,$6,0;a,$10,0,$5;a,$11,$5,0;a,$1,0,cd4;g,hd2,$14;h,$15,hd2,0;m,$16,$5,hd2;m,$17,1,2;a,hc,$15,$18;m,$15,$12,hd2;a,$19,0,$20;m,$15,$10,hd2;a,$19,0,$22;m,$15,$11,hd2;a,$19,0,$24;m,$15,$13,hd2;a,$19,0,$26;m,$15,2,1;a,$19,0,$28;a,$19,$17,0;a,$21,$17,0;a,$23,$17,0;a,$25,$17,0;a,$27,$17,0;a,$29,$17,0;m,$5,hd2,$16;a,$19,$36,0;a,$30,0,$36;q,$14,$37,$30;q,$14,$19,$38;m,$5,$15,$16;a,0,0,$41;q,$14,$41,t;q,$14,t,$42;a,w,0,$39;a,w,0,$40;a,h,0,$43;a,h,0,$44;d,$39,$40,2;d,$46,$45,2;d,$43,$44,2;d,$12,$10,2;d,$11,$13,2;d,$48,$47,2|f,1,0,0:M,$8,$12;L,$21,$12;L,$40,$44;L,$39,$43;L,$31,$12;L,$9,$12;L,$9,$10;L,$32,$10;L,$33,$11;L,$9,$11;L,$9,$13;L,$34,$13;L,$46,$48;L,$45,$47;L,$27,$13;L,$8,$13;L,$8,$11;L,$25,$11;L,$23,$10;L,$8,$10;Z|$8,$12,$9,$1"
            s = s & "3"
        Case 169      ' cornerTabs
            s = "|h,w,h,0;m,1,$0,20;a,0,b,$1;a,0,r,$1|f,1,0,0:M,l,t;L,$1,t;L,l,$1;Z~f,1,0,0:M,l,$2;L,$1,b;L,l,b;Z~f,1,0,0:M,$3,t;L,r,t;L,r,$1;Z~f,1,0,0:M,r,$2;L,r,b;L,$3,b;Z|$1,$1,$3,$2"
        Case 170      ' squareTabs
            s = "|h,w,h,0;m,1,$0,20;a,0,b,$1;a,0,r,$1|f,1,0,0:M,l,t;L,$1,t;L,$1,$1;L,l,$1;Z~f,1,0,0:M,l,$2;L,$1,$2;L,$1,b;L,l,b;Z~f,1,0,0:M,$3,t;L,r,t;L,r,$1;L,$3,$1;Z~f,1,0,0:M,$3,$2;L,r,$2;L,r,b;L,$3,b;Z|$1,$1,$3,$2"
        Case 171      ' plaqueTabs
            s = "|h,w,h,0;m,1,$0,20;a,0,b,$1;a,0,r,$1|f,1,0,0:M,l,t;L,$1,t;A,$1,$1,0,cd4;Z~f,1,0,0:M,l,$2;A,$1,$1,k34,cd4;L,l,b;Z~f,1,0,0:M,r,t;L,r,$1;A,$1,$1,cd4,cd4;Z~f,1,0,0:M,$3,b;A,$1,$1,cd2,cd4;L,r,b;Z|$1,$1,$3,$2"
        Case 172      ' gear6
            s = "15000,3526|p,0,@1,20000;p,0,@2,5358;m,ss,$0,100000;m,ss,$1,100000;m,$2,1,2;m,$3,1,2;a,$4,$5,0;a,hd2,0,$2;a,wd2,0,$2;a,$8,0,$7;q,$9,$7,$8;t,$10,$6;a,19800000,0,$11;a,19800000,$11,0;o,$8,$12;i,$7,$12;t,$14,$15;o,$7,$16;i,$8,$16;h,$17,$18,0;m,$8,$7,$19;o,$20,$16;i,$20,$16;a,hc,$21,0;a,vc,$22,0;o,$8,$13;i,$7,$13;t,$25,$26;o,$7,$27;i,$8,$27;h,$28,$29,0;m,$8,$7,$30;o,$31,$27;i,$31,$27;a,hc,$32,0;a,vc,$33,0;a,$23,0,$34;a,$24,0,$35;h,$36,$37,0;t,$37,$36;i,$3,$39;o,$3,$39;a,$34,$40,0;a,$35,$41,0;a,$23,0,$40;a,$24,0,$41;i,$2,$39;o,$2,$39;a,$43,$46,0;a,$42,0,$47;a,$45,$46,0;a,$44,0,$47;a,k34,$11,0;o,$8,$52;i,$7,$52;t,$53,$54;o,$7,$55;i,$8,$55;h,$56,$57,0;m,$8,$7,$58;o,$59,$55;i,$59,$55;a,hc,$60,0;a,vc,$61,0;a,hc,0,$60;a,$62,0,$3;a,$64,$3,0;a,$63,0,$2;a,$16,0,$55;a,1800000,0,$11;a,1800000,$11,0;o,$8,$69;i,$7,$69;t,$71,$72;a,h,0,$35;o,$8,$70;i,$7,$70;t,$75,$76;a,h,0,$24;a,h,0,$50;a,h,0,$48;v,$49;a,$7"
            s = s & "3,0,$27;a,cd4,$11,0;o,$8,$83;i,$7,$83;t,$84,$85;a,h,0,$63;a,h,0,$67;a,9000000,$11,0;o,$8,$89;i,$7,$89;t,$90,$91;a,w,0,$34;a,w,0,$49;a,w,0,$51;a,12600000,$11,0;o,$8,$96;i,$7,$96;t,$97,$98;a,w,0,$23;a,w,0,$51;a,w,0,$49;d,$51,$49,2;d,$50,$48,2;a,b,0,$104;d,r,0,$103|f,1,0,0:M,$23,$24;L,$51,$50;L,$49,$48;L,$34,$35;A,$8,$7,$27,$82;L,$49,$80;L,$51,$79;L,$23,$78;A,$8,$7,$77,$68;L,$65,$88;L,$66,$88;L,$64,$87;A,$8,$7,$86,$68;L,$95,$79;L,$94,$80;L,$93,$74;A,$8,$7,$92,$82;L,$102,$48;L,$101,$50;L,$100,$24;A,$8,$7,$99,$68;L,$66,$67;L,$65,$67;L,$62,$63;A,$8,$7,$55,$68;Z|$100,$24,$23,$78"
        Case 173      ' gear9
            s = "10000,1763|p,0,@1,20000;p,0,@2,2679;m,ss,$0,100000;m,ss,$1,100000;m,$2,1,2;m,$3,1,2;a,$4,$5,0;a,hd2,0,$2;a,wd2,0,$2;a,$8,0,$7;q,$9,$7,$8;t,$10,$6;a,18600000,0,$11;a,18600000,$11,0;o,$8,$12;i,$7,$12;t,$14,$15;o,$7,$16;i,$8,$16;h,$17,$18,0;m,$8,$7,$19;o,$20,$16;i,$20,$16;a,hc,$21,0;a,vc,$22,0;o,$8,$13;i,$7,$13;t,$25,$26;o,$7,$27;i,$8,$27;h,$28,$29,0;m,$8,$7,$30;o,$31,$27;i,$31,$27;a,hc,$32,0;a,vc,$33,0;a,$23,0,$34;a,$24,0,$35;h,$36,$37,0;t,$37,$36;i,$3,$39;o,$3,$39;a,$34,$40,0;a,$35,$41,0;a,$23,0,$40;a,$24,0,$41;i,$2,$39;o,$2,$39;a,$43,$46,0;a,$42,0,$47;a,$45,$46,0;a,$44,0,$47;a,21000000,0,$11;a,21000000,$11,0;o,$8,$52;i,$7,$52;t,$54,$55;o,$7,$56;i,$8,$56;h,$57,$58,0;m,$8,$7,$59;o,$60,$56;i,$60,$56;a,hc,$61,0;a,vc,$62,0;o,$8,$53;i,$7,$53;t,$65,$66;o,$7,$67;i,$8,$67;h,$68,$69,0;m,$8,$7,$70;o,$71,$67;i,$71,$67;a,hc,$72,0;a,vc,$73,0;a,$63,0,$74;a,$64,0,$75;h,$76,$77,0;t,$77,$76;i,$3,$79;o,$3,"
            s = s & "$79;a,$74,$80,0;a,$75,$81,0;a,$63,0,$80;a,$64,0,$81;i,$2,$79;o,$2,$79;a,$83,$86,0;a,$82,0,$87;a,$85,$86,0;a,$84,0,$87;a,$56,0,$27;a,1800000,0,$11;a,1800000,$11,0;o,$8,$93;i,$7,$93;t,$95,$96;o,$7,$97;i,$8,$97;h,$98,$99,0;m,$8,$7,$100;o,$101,$97;i,$101,$97;a,hc,$102,0;a,vc,$103,0;o,$8,$94;i,$7,$94;t,$106,$107;o,$7,$108;i,$8,$108;h,$109,$110,0;m,$8,$7,$111;o,$112,$108;i,$112,$108;a,hc,$113,0;a,vc,$114,0;a,$104,0,$115;a,$105,0,$116;h,$117,$118,0;t,$118,$117;i,$3,$120;o,$3,$120;a,$115,$121,0;a,$116,$122,0;a,$104,0,$121;a,$105,0,$122;i,$2,$120;o,$2,$120;a,$124,$127,0;a,$123,0,$128;a,$126,$127,0;a,$125,0,$128;a,$97,0,$67;a,4200000,0,$11;a,4200000,$11,0;o,$8,$134;i,$7,$134;t,$136,$137;o,$7,$138;i,$8,$138;h,$139,$140,0;m,$8,$7,$141;o,$142,$138;i,$142,$138;a,hc,$143,0;a,vc,$144,0;o,$8,$135;i,$7,$135;t,$147,$148;o,$7,$149;i,$8,$149;h,$150,$151,0;m,$8,$7,$152;o,$153,$149;i,$153,$149;a,hc,$154,0;a,vc"
            s = s & ",$155,0;a,$145,0,$156;a,$146,0,$157;h,$158,$159,0;t,$159,$158;i,$3,$161;o,$3,$161;a,$156,$162,0;a,$157,$163,0;a,$145,0,$162;a,$146,0,$163;i,$2,$161;o,$2,$161;a,$165,$168,0;a,$164,0,$169;a,$167,$168,0;a,$166,0,$169;a,$138,0,$108;a,6600000,0,$11;a,6600000,$11,0;o,$8,$175;i,$7,$175;t,$177,$178;o,$8,$176;i,$7,$176;t,$180,$181;a,w,0,$145;a,w,0,$173;a,w,0,$171;a,$179,0,$149;a,9000000,$11,0;o,$8,$187;i,$7,$187;t,$188,$189;a,w,0,$104;a,w,0,$132;a,w,0,$130;a,11400000,$11,0;o,$8,$194;i,$7,$194;t,$195,$196;a,w,0,$63;a,w,0,$91;a,w,0,$89;a,13800000,$11,0;o,$8,$201;i,$7,$201;t,$202,$203;a,w,0,$34;a,w,0,$23;a,w,0,$51;a,w,0,$49;a,k34,0,$11;a,k34,$11,0;o,$8,$210;i,$7,$210;t,$211,$212;o,$7,$213;i,$8,$213;h,$214,$215,0;m,$8,$7,$216;o,$217,$213;i,$217,$213;a,hc,$218,0;a,vc,$219,0;o,$8,$209;i,$7,$209;t,$222,$223;a,hc,0,$218;a,$220,0,$3;a,$225,$3,0;a,$221,0,$2;a,$224,0,$204;d,$51,$49,2;d,$50,$48,2;d,$91,$89,2"
            s = s & ";d,$90,$88,2;d,$132,$130,2;d,$131,$129,2;d,$173,$171,2;d,$172,$170,2;d,r,0,$236;d,r,0,$234;d,r,0,$232;d,r,0,$230|f,1,0,0:M,$23,$24;L,$51,$50;L,$49,$48;L,$34,$35;A,$8,$7,$27,$92;L,$91,$90;L,$89,$88;L,$74,$75;A,$8,$7,$67,$133;L,$132,$131;L,$130,$129;L,$115,$116;A,$8,$7,$108,$174;L,$173,$172;L,$171,$170;L,$156,$157;A,$8,$7,$149,$186;L,$185,$170;L,$184,$172;L,$183,$146;A,$8,$7,$182,$174;L,$193,$129;L,$192,$131;L,$191,$105;A,$8,$7,$190,$133;L,$200,$88;L,$199,$90;L,$198,$64;A,$8,$7,$197,$92;L,$208,$48;L,$207,$50;L,$206,$24;A,$8,$7,$204,$229;L,$227,$228;L,$226,$228;L,$220,$221;A,$8,$7,$213,$229;Z|$205,$35,$34,$116"
        Case 174      ' funnel
            s = "|m,ss,1,20;a,wd2,0,$0;a,hd4,0,$0;o,wd2,480000;i,hd4,480000;t,$3,$4;m,$5,2,1;a,cd2,0,$5;a,cd2,$6,0;a,cd2,0,$6;m,wd2,1,4;m,hd4,1,4;o,hd4,$7;i,wd2,$7;h,$12,$13,0;m,wd2,hd4,$14;o,$15,$7;i,$15,$7;a,hc,$16,0;a,hd4,$17,0;o,$11,$5;i,$10,$5;h,$20,$21,0;m,$10,$11,$22;o,$23,$5;i,$23,$5;a,hc,$24,0;a,b,0,$11;a,$27,$25,0;a,wd2,0,$1;m,cd2,2,1|f,1,0,0:M,$18,$19;A,wd2,hd4,$7,$8;L,$26,$28;A,$10,$11,$5,$9;Z;M,$29,hd4;A,$1,$2,cd2,-21600000;Z|l,t,r,b"
        Case 175      ' pieWedge
            s = "|o,w,13500000;i,h,13500000;a,r,$0,0;a,b,$1,0|f,1,0,0:M,l,b;A,w,h,cd2,cd4;L,r,b;Z|$2,$3,r,b"
        Case 176      ' leftCircularArrow
            s = "12500,!-1142319,!1142319,!10800000,12500|p,0,@5,25000;m,$0,2,1;p,0,@1,$1;p,1,@3,21599999;p,0,@4,21599999;m,ss,$2,100000;m,ss,$0,100000;m,$5,1,2;a,wd2,$7,$6;a,hd2,$7,$6;a,$8,0,$5;a,$9,0,$5;a,$10,$7,0;a,$11,$7,0;i,$12,$3;o,$13,$3;c,$12,$15,$14;s,$13,$15,$14;a,hc,$16,0;a,vc,$17,0;n,$10,$11;m,$16,$16,1;m,$17,$17,1;m,$20,$20,1;a,$21,0,$23;a,$22,0,$23;m,$24,$25,$21;m,$26,1,$22;a,1,0,$27;r,$28;m,$24,1,$16;m,$30,1,$17;d,1,$29,$31;t,1,$32;a,$33,21600000,0;q,$33,$33,$34;a,$35,0,$3;a,$36,21600000,0;q,$36,$36,$37;a,$38,0,cd2;a,$38,0,21600000;q,$39,$40,$38;b,$41;m,$42,-1,1;b,@2;m,$44,-1,1;p,$43,$45,0;a,$3,$46,0;i,$12,$47;o,$13,$47;c,$12,$49,$48;s,$13,$49,$48;a,hc,$50,0;a,vc,$51,0;i,$8,$4;o,$9,$4;c,$8,$55,$54;s,$9,$55,$54;a,hc,$56,0;a,vc,$57,0;i,$10,$4;o,$11,$4;c,$10,$61,$60;s,$11,$61,$60;a,hc,$62,0;a,vc,$63,0;o,$6,$47;i,$6,$47;a,$18,$66,0;a,$19,$67,0;o,$6,$47;i,$6,$47;a,$18,0,$70,0;a,$19,0,$71,0;a,$7"
            s = s & "2,0,hc;a,$73,0,vc;a,$68,0,hc;a,$69,0,vc;n,$8,$9;m,$74,$78,$8;m,$75,$78,$9;m,$76,$78,$8;m,$77,$78,$9;a,$81,0,$79;a,$82,0,$80;h,$83,$84,0;m,$79,$82,1;m,$81,$80,1;a,$86,0,$87;m,$78,$78,1;m,$85,$85,1;m,$89,$90,1;m,$88,$88,1;a,$91,0,$92;x,$93,0;r,$94;m,$84,-1,1;q,$96,-1,1;m,$97,$83,1;m,$98,$95,1;m,$88,$84,1;d,$100,$99,$90;a,$100,0,$99;m,$102,1,$90;b,$84;m,$104,$95,1;m,$88,$83,-1;d,$106,$105,$90;a,$106,0,$105;m,$108,1,$90;a,$81,0,$101;a,$81,0,$103;a,$82,0,$107;a,$82,0,$109;h,$110,$112,0;h,$111,$113,0;a,$115,0,$114;q,$116,$101,$103;q,$116,$107,$109;m,$117,$8,$78;m,$118,$9,$78;a,hc,$119,0;a,vc,$120,0;m,$74,$20,$10;m,$75,$20,$11;m,$76,$20,$10;m,$77,$20,$11;a,$125,0,$123;a,$126,0,$124;h,$127,$128,0;m,$123,$126,1;m,$125,$124,1;a,$130,0,$131;m,$20,$20,1;m,$129,$129,1;m,$133,$134,1;m,$132,$132,1;a,$135,0,$136;x,$137,0;r,$138;m,$97,$127,1;m,$140,$139,1;m,$132,$128,1;d,$142,$141,$134;a,$142,0,$141;m,$1"
            s = s & "44,1,$134;b,$128;m,$146,$139,1;m,$132,$127,-1;d,$148,$147,$134;a,$148,0,$147;m,$150,1,$134;a,$123,0,$143;a,$123,0,$145;a,$124,0,$149;a,$124,0,$151;h,$152,$154,0;h,$153,$155,0;a,$157,0,$156;q,$158,$143,$145;q,$158,$149,$151;m,$159,$10,$20;m,$160,$11,$20;a,hc,$161,0;a,vc,$162,0;t,$161,$162;a,$165,21600000,0;q,$165,$165,$166;a,$4,0,$167;a,$168,21600000,0;q,$168,$168,$169;a,$167,$170,0;a,0,0,$170;a,$121,0,$163;a,$122,0,$164;h,$173,$174,0;m,$175,1,2;a,$176,0,$6;q,$177,$121,$68;q,$177,$122,$69;q,$177,$163,$72;q,$177,$164,$73;t,$119,$120;a,$182,21600000,0;q,$182,$182,$183;a,$184,0,$4;a,$185,0,21600000;q,$185,$186,$185;a,$4,$187,0;a,0,0,$187;i,$12,$4;o,$13,$4;c,$12,$191,$190;s,$13,$191,$190;a,hc,$192,0;a,vc,$193,0;a,$4,cd4,0;a,$47,0,cd4;a,$47,cd2,0;o,$8,2700000;i,$9,2700000;a,hc,0,$199;a,hc,$199,0;a,vc,0,$200;a,vc,$200,0|f,1,0,0:M,$58,$59;L,$64,$65;A,$10,$11,$171,$172;L,$180,$181;L,$52,$53;L,$17"
            s = s & "8,$179;L,$121,$122;A,$8,$9,$188,$189;Z|$201,$203,$202,$204"
        Case 177      ' leftRightCircularArrow
            s = "12500,!1142319,!20457681,!11942319,12500|p,0,@5,25000;m,$0,2,1;p,0,@1,$1;p,1,@3,21599999;p,0,@4,21599999;m,ss,$2,100000;m,ss,$0,100000;m,$5,1,2;a,wd2,$7,$6;a,hd2,$7,$6;a,$8,0,$5;a,$9,0,$5;a,$10,$7,0;a,$11,$7,0;i,$12,$3;o,$13,$3;c,$12,$15,$14;s,$13,$15,$14;a,hc,$16,0;a,vc,$17,0;n,$10,$11;m,$16,$16,1;m,$17,$17,1;m,$20,$20,1;a,$21,0,$23;a,$22,0,$23;m,$24,$25,$21;m,$26,1,$22;a,1,0,$27;r,$28;m,$24,1,$16;m,$30,1,$17;d,1,$29,$31;t,1,$32;a,$33,21600000,0;q,$33,$33,$34;a,$35,0,$3;a,$36,21600000,0;q,$36,$36,$37;a,$38,0,cd2;a,$38,0,21600000;q,$39,$40,$38;b,$41;p,0,@2,$42;a,$3,$43,0;i,$12,$44;o,$13,$44;c,$12,$46,$45;s,$13,$46,$45;a,hc,$47,0;a,vc,$48,0;o,$6,$44;i,$6,$44;a,$18,$51,0;a,$19,$52,0;o,$6,$44;i,$6,$44;a,$18,0,$55,0;a,$19,0,$56,0;a,$57,0,hc;a,$58,0,vc;a,$53,0,hc;a,$54,0,vc;n,$8,$9;m,$59,$63,$8;m,$60,$63,$9;m,$61,$63,$8;m,$62,$63,$9;a,$66,0,$64;a,$67,0,$65;h,$68,$69,0;m,$64,$67,1;m,$66,$65,1;"
            s = s & "a,$71,0,$72;m,$63,$63,1;m,$70,$70,1;m,$74,$75,1;m,$73,$73,1;a,$76,0,$77;x,$78,0;r,$79;m,$69,-1,1;q,$81,-1,1;m,$82,$68,1;m,$83,$80,1;m,$73,$69,1;d,$85,$84,$75;a,$85,0,$84;m,$87,1,$75;b,$69;m,$89,$80,1;m,$73,$68,-1;d,$91,$90,$75;a,$91,0,$90;m,$93,1,$75;a,$66,0,$86;a,$66,0,$88;a,$67,0,$92;a,$67,0,$94;h,$95,$97,0;h,$96,$98,0;a,$100,0,$99;q,$101,$86,$88;q,$101,$92,$94;m,$102,$8,$63;m,$103,$9,$63;a,hc,$104,0;a,vc,$105,0;m,$59,$20,$10;m,$60,$20,$11;m,$61,$20,$10;m,$62,$20,$11;a,$110,0,$108;a,$111,0,$109;h,$112,$113,0;m,$108,$111,1;m,$110,$109,1;a,$115,0,$116;m,$20,$20,1;m,$114,$114,1;m,$118,$119,1;m,$117,$117,1;a,$120,0,$121;x,$122,0;r,$123;m,$82,$112,1;m,$125,$124,1;m,$117,$113,1;d,$127,$126,$119;a,$127,0,$126;m,$129,1,$119;b,$113;m,$131,$124,1;m,$117,$112,-1;d,$133,$132,$119;a,$133,0,$132;m,$135,1,$119;a,$108,0,$128;a,$108,0,$130;a,$109,0,$134;a,$109,0,$136;h,$137,$139,0;h,$138,$140,0;a,$142,"
            s = s & "0,$141;q,$143,$128,$130;q,$143,$134,$136;m,$144,$10,$20;m,$145,$11,$20;a,hc,$146,0;a,vc,$147,0;i,$12,$4;o,$13,$4;c,$12,$151,$150;s,$13,$151,$150;a,hc,$152,0;a,vc,$153,0;a,$4,0,$43;i,$12,$156;o,$13,$156;c,$12,$158,$157;s,$13,$158,$157;a,hc,$159,0;a,vc,$160,0;o,$6,$156;i,$6,$156;a,$154,$163,0;a,$155,$164,0;o,$6,$156;i,$6,$156;a,$154,0,$167,0;a,$155,0,$168,0;a,$106,0,$148;a,$107,0,$149;h,$171,$172,0;m,$173,1,2;a,$174,0,$6;q,$175,$106,$53;q,$175,$107,$54;q,$175,$148,$57;q,$175,$149,$58;t,$104,$105;a,$180,21600000,0;q,$180,$180,$181;a,$182,0,$3;a,$183,21600000,0;q,$183,$183,$184;a,$4,0,$185;a,$186,21600000,0;q,$186,$186,$187;a,$182,0,$188;a,$189,21600000,0;q,$189,$189,$190;t,$146,$147;a,$192,21600000,0;q,$192,$192,$193;a,$194,0,$3;a,$195,0,21600000;q,$195,$196,$195;a,$4,0,$197;a,$198,0,21600000;q,$199,$199,$198;a,$200,0,$194;a,$201,0,21600000;q,$201,$202,$201;i,$8,$188;o,$9,$188;c,$8,$205,$20"
            s = s & "4;s,$9,$205,$204;a,hc,$206,0;a,vc,$207,0;i,$10,$200;o,$11,$200;c,$10,$211,$210;s,$11,$211,$210;a,hc,$212,0;a,vc,$213,0;q,$175,$208,$165;q,$175,$209,$166;q,$175,$214,$169;q,$175,$215,$170;a,$156,0,cd4;a,$44,cd4,0;a,$44,cd2,0;a,$156,cd2,0;o,$8,2700000;i,$9,2700000;a,hc,0,$224;a,hc,$224,0;a,vc,0,$225;a,vc,$225,0|f,1,0,0:M,$161,$162;L,$216,$217;L,$208,$209;A,$8,$9,$188,$191;L,$176,$177;L,$49,$50;L,$178,$179;L,$148,$149;A,$10,$11,$194,$203;L,$218,$219;Z|$226,$228,$227,$229"
        Case 178      ' swooshArrow
            s = "25000,16667|p,1,@1,75000;m,70000,w,ss;p,0,@2,$1;m,h,$0,100000;m,ss,$2,100000;a,r,0,$4;a,t,ssd8,0;m,cd4,1,14;g,ssd8,$7;a,$5,0,$8;g,$3,$7;a,$6,$3,0;a,$5,$10,0;a,$12,$8,0;a,$11,ssd8,0;a,$14,0,t;m,$15,1,2;m,h,1,20;a,t,$16,$17;m,hd6,1,1;a,hd6,$19,0;v,wd6;m,hd6,1,2;a,$11,$22,0;v,wd4|f,1,0,0:M,l,b;Q,$21,$20,$5,$6;L,$9,t;L,r,$18;L,$13,$14;L,$12,$11;Q,$24,$23,l,b;Z|l,t,r,b"
        Case 179      ' cloud
            s = "|m,w,2977,21600;m,h,3262,21600;m,w,17087,21600;m,h,17337,21600;m,w,67,21600;m,h,21577,21600;m,w,21582,21600;m,h,1235,21600|f,1,43200,43200:M,3900,14370;A,6753,9190,-11429249,7426832;A,5333,7267,-8646143,5396714;A,4365,5945,-8748475,5983381;A,4857,6595,-7859164,7034504;A,5333,7273,-4722533,6541615;A,6775,9220,-2776035,7816140;A,5785,7867,37501,6842000;A,6752,9215,1347096,6910353;A,7720,10543,3974558,4542661;A,4360,5918,-16496525,8804134;A,4345,5945,-14809710,9151131;Z~n,1,43200,43200:M,4693,26177;A,4345,5945,5204520,1585770;M,6928,34899;A,4360,5918,4416628,686848;M,16478,39090;A,6752,9215,8257449,844866;M,28827,34751;A,6752,9215,387196,959901;M,34129,22954;A,5785,7867,-4217541,4255042;M,41798,15354;A,5333,7273,1819082,1665090;M,38324,5426;A,4857,6595,-824660,891534;M,29078,3952;A,4857,6595,-8950887,1091722;M,22141,4720;A,4365,5945,-9809656,1061181;M,14000,5192;A,6753,9190,-4002417,739161;"
            s = s & "M,4127,15789;A,6753,9190,9459261,711490|$0,$1,$2,$3"
        Case 180      ' chartX
            s = "||n,1,10,10:M,0,0;L,10,10;M,0,10;L,10,0~f,0,10,10:M,0,0;L,0,10;L,10,10;L,10,0;Z|l,t,r,b"
        Case 181      ' chartStar
            s = "||n,1,10,10:M,0,0;L,10,10;M,0,10;L,10,0;M,5,0;L,5,10~f,0,10,10:M,0,0;L,0,10;L,10,10;L,10,0;Z|l,t,r,b"
        Case 182      ' chartPlus
            s = "||n,1,10,10:M,5,0;L,5,10;M,0,5;L,10,5~f,0,10,10:M,0,0;L,0,10;L,10,10;L,10,0;Z|l,t,r,b"
        Case 183      ' lineInv
            s = "||f,1,0,0:M,l,b;L,r,t|l,t,r,b"
    End Select
    PresetSrc3 = s
End Function

Private Function PresetSource(ByVal ast As Long) As String
    Dim s As String
    If Len(s) = 0 Then s = PresetSrc1(ast)
    If Len(s) = 0 Then s = PresetSrc2(ast)
    If Len(s) = 0 Then s = PresetSrc3(ast)
    PresetSource = s
End Function


Private Sub LoadStrings(ByVal lang As String)
    Select Case lang
        Case "zh-TW"
            LoadStrings_zh_TW_1
            LoadStrings_zh_TW_2
            LoadStrings_zh_TW_3
            LoadStrings_zh_TW_4
            LoadStrings_zh_TW_5
            LoadStrings_zh_TW_6
        Case "en"
            LoadStrings_en_1
            LoadStrings_en_2
            LoadStrings_en_3
            LoadStrings_en_4
            LoadStrings_en_5
            LoadStrings_en_6
        Case "ja"
            LoadStrings_ja_1
            LoadStrings_ja_2
            LoadStrings_ja_3
            LoadStrings_ja_4
            LoadStrings_ja_5
            LoadStrings_ja_6
        Case Else
            LoadStrings_zh_CN_1
            LoadStrings_zh_CN_2
            LoadStrings_zh_CN_3
            LoadStrings_zh_CN_4
            LoadStrings_zh_CN_5
            LoadStrings_zh_CN_6
    End Select
End Sub

Private Sub LoadStrings_zh_CN_1()
    Dim s As String
    gStr.Add U("754C97628BED8A00FF1A007B0030007DFF0853EF5728529F80FD533A201C51734E8E201D7EC476848BED8A0083DC53554E2D52076362FF09"), "about.lang"
    s = ""
    s = s & U("77E253700020005600650063005300740061006D0070002000200076007B0030007D")
    s = s & vbCrLf
    s = s & U("90094E2D53735370002020142014002056FE5F624E00952E5BFC51FA00200045004D00460020002F00200053005600470020002F00200050004400460020002F002000560053004400580020002F00200050004E004700207B49683C5F0F")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("4F5C8005FF1A007B0031007D")
    s = s & vbCrLf
    s = s & U("90AE7BB1FF1A007B0032007D")
    s = s & vbCrLf
    s = s & U("4E3B9875FF1A007B0033007D")
    s = s & vbCrLf
    s = s & U("5F006E90534F8BAEFF1A004D004900540020004C006900630065006E00730065")
    s = s & vbCrLf
    s = s & "Copyright (c) {4} {1}"
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("77E25370627F8BFA6C384E455F006E9030016C384E45514D8D39FF1A6CA1670965368D39529F80FDFF0C6CA167095E7F544AFF0C4E0D80547F51653696C64EFB4F554FE1606FFF0C516890E86E907801572800200047006900740048007500620020516C5F003002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("5982679C5B834E3A4F60828277014E8665F695F4FF0C6B228FCE70B951FB529F80FD533A201C51734E8E201D7EC47684201C8D5E8D4F4F5C8005201DFF0C75285FAE4FE16216652F4ED85B9D626B78018BF74F5C8005559D676F5496556130028D5E8D4F5B8C516881EA613FFF0C4E0D5F7154CD4EFB4F55529F80FDFF1B4F607684652F63014F1A75284E8E63017EED7EF462A4548C65398FDB77E253703002")
    gStr.Add s, "about.text"
    gStr.Add U("51734E8E0020007B0030007D"), "about.title"
    gStr.Add U("77E253700020005600650063005300740061006D0070"), "app.title"
    gStr.Add U("7C738272"), "bg.beige"
    gStr.Add U("81EA5B9A4E49989C8272"), "bg.custom"
    gStr.Add U("81EA5B9A4E49989C82720020007B0030007D"), "bg.customValue"
    gStr.Add U("7F51683C"), "bg.grid"
    gStr.Add U("7F51683CFF0895F48DDD0020007B0030007D002078C5FF09"), "bg.gridValue"
    gStr.Add U("900F660E"), "bg.none"
    gStr.Add U("767D8272"), "bg.white"
    gStr.Add U("53D66D88"), "btn.cancel"
    gStr.Add U("786E5B9A"), "btn.ok"
    gStr.Add U("989C8272683C5F0F793A4F8BFF1A003200350035002C003200340038002C0032003300350020621600200023004600460046003800450042"), "col.bad"
    gStr.Add U("6DF182720020007B0030007D0025"), "col.darker"
    gStr.Add U("535551FB8272575790098272FF0C53CC51FB76F463A5786E5B9AFF1B4E5F53EF8F93516500200052002C0047002C004200206216002000230052005200470047004200423002"), "col.hint"
    gStr.Add U("6DE182720020007B0030007D0025"), "col.lighter"
    gStr.Add U("8F9351657F1653F7900962E9989C8272FF0C62168F93516500200052002C0047002C00420020002F00200023005200520047004700420042FF1A"), "col.listPrompt"
    gStr.Add U("67008FD14F7F75287684989C8272"), "col.recent"
    gStr.Add U("67008FD14F7F75287684989C8272FF1A007B0030007D"), "col.recentItem"
    gStr.Add U("680751C68272"), "col.standard"
    gStr.Add U("680751C68272FF1A007B0030007D"), "col.stdItem"
    gStr.Add U("6DF17EA2007C7EA28272007C6A598272007C9EC48272007C6D457EFF007C7EFF8272007C6D4584DD007C84DD8272007C6DF184DD007C7D2B8272"), "col.stdNames"
    gStr.Add U("7CFB7EDF989C82722026"), "col.system"
    gStr.Add U("7CFB7EDF989C8272"), "col.systemPicked"
    gStr.Add U("4E3B9898989C8272"), "col.theme"
    gStr.Add U("767D8272FF0C80CC666F00200031007C9ED18272FF0C65875B5700200031007C80CC666F00200032007C65875B5700200032007C4E2A6027827200200031007C4E2A6027827200200032007C4E2A6027827200200033007C4E2A6027827200200034007C4E2A6027827200200035007C4E2A6027827200200036"), "col.themeNames"
    gStr.Add U("007B0030007DFF0C007B0031007D"), "col.tint"
    gStr.Add U("81EA5B9A4E49989C8272"), "col.title"
    gStr.Add U("989C8272503C"), "col.value"
    gStr.Add U("007B0030007DFF08007B0031007DFF09"), "col.withValue"
    gStr.Add U("65874EF6540DFF084E0D542B62695C55540DFF09FF1A"), "dlg.fileName"
    gStr.Add U("900962E95BFC51FA65874EF65939"), "dlg.pickFolder"
    gStr.Add U("900962E9901098755BFC51FA768465874EF65939FF08683C5F0FFF1A007B0030007DFF09"), "dlg.pickFolderBatch"
    gStr.Add U("900962E95BFC51FA65874EF65939FF086BCF4E2A56FE5F624E004E2A65874EF6FF09"), "dlg.pickFolderEach"
    gStr.Add U("5BFC51FA4E3A0020007B0030007D"), "dlg.saveAs"
    gStr.Add U("5BFC51FA90094E2D56FE5F62"), "dlg.saveTitle"
    gStr.Add U("652F4ED85B9D"), "don.alipay"
    s = ""
    s = s & U("8C228C224F60613F610F652F630177E25370FF01")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("8BF77528007B0030007D626B63CF201C8D5E8D4F4F5C8005201D4E2D5BF95E9476844E8C7EF47801FF0C91D1989D968F610F3002")
    s = s & vbCrLf
    s = s & U("8D5E8D4F5B8C516881EA613FFF0C77E25370768462406709529F80FD90FD6C384E45514D8D3930015F006E903002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("5982679C65B94FBFFF0C4E5F6B228FCE5728002000470069007400480075006200207ED9987976EE70B94E004E2A00200053007400610072FF0C6216628A5B8363A883507ED9970089817684540C5B66548C540C4E8B3002")
    gStr.Add s, "don.thanks"
    gStr.Add U("8D5E8D4F"), "don.title"
    gStr.Add U("5FAE4FE1"), "don.wechat"
    gStr.Add U("81EA5B9A4E49"), "dpi.custom"
    gStr.Add U("65E06CD5901A8FC7526A8D34677F590D52366240900956FE5F62FF08526A8D34677F53EF80FD88AB51764ED67A0B5E8F53607528FF09FF0C8BF791CD8BD53002"), "err.clipboard"
    gStr.Add U("0045004D005A0020538B7F2959318D253002"), "err.emz"
    gStr.Add U("624090095BF98C6165E06CD57EC454085BFC51FA3002"), "err.group"
    gStr.Add U("627E4E0D52300020006D00610063004F005300208F8552A9811A672C0020007B0030007DFF0C8BF791CD65B08FD0884C0020006D00610063004F005300205B8988C556683002"), "err.macHelper"
    gStr.Add U("FF080050006F0077006500720050006F0069006E007400200066006F00720020004D00610063002097008981002000310036002E003800320020621666F49AD87248672C624D80FD5BFC51FA56FE5F62FF09"), "err.macVersion"
    gStr.Add U("65E06CD5628A65874EF651995165FF1A007B0030007D"), "err.macWrite"
    gStr.Add U("6CA16709751F621065874EF6FF1A007B0030007D"), "err.noFile"
    gStr.Add U("00500044004600205BFC51FA59318D25FF1A007B0030007D"), "err.pdf"
End Sub

Private Sub LoadStrings_zh_CN_2()
    Dim s As String
    gStr.Add U("6253530500200056005300440058002059318D253002"), "err.zip"
    gStr.Add U("6253530500200056005300440058002059318D25FF08970089817CFB7EDF81EA5E267684002000570069006E0064006F0077007300200050006F007700650072005300680065006C006CFF093002"), "err.zipPs"
    s = ""
    s = s & U("8BF7900962E953CD998865B95F0FFF1A")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("3010662F301153D1900190AE4EF67ED94F5C8005FF08007B0030007DFF09")
    s = s & vbCrLf
    s = s & U("30105426301157280020004700690074004800750062002063D04EA4002000490073007300750065FF0897008981002000470069007400480075006200208D2653F7FF0C95EE98984F1A516C5F00663E793AFF09")
    s = s & vbCrLf
    s = s & U("301053D66D8830118FD456DE")
    gStr.Add s, "fb.ask"
    s = ""
    s = s & U("002A002A95EE989863CF8FF00020002F00205EFA8BAEFF1A002A002A")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("002A002A590D73B06B659AA4FF1A002A002A")
    s = s & vbCrLf
    s = s & "1. "
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("002A002A8FD0884C73AF5883FF1A002A002A")
    s = s & vbCrLf
    s = s & "```"
    s = s & vbCrLf
    s = s & "{0}"
    s = s & vbCrLf
    s = s & "```"
    s = s & vbCrLf
    gStr.Add s, "fb.issueBody"
    gStr.Add U("005B53CD9988005D0020"), "fb.issueTitle"
    s = ""
    s = s & U("4F5C800590AE7BB15DF2590D52365230526A8D34677FFF1A007B0030007D")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("4F6053EF4EE557284EFB4F5590AE7BB1FF087F51987590AE7BB1300100510051002090AE7BB13001004F00750074006C006F006F006B00207B49FF094E2D65B05EFA90AE4EF630017C988D3465364EF64EBA540E53D190013002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("898173B0572862535F00753581114E0A76849ED88BA490AE4EF67A0B5E8F5417FF1F")
    s = s & vbCrLf
    s = s & U("FF085982679C6CA167098BBE7F6E90AE4EF67A0B5E8FFF0C53EF4EE5900962E9201C5426201DFF0C76F463A575287F51987590AE7BB153D19001FF09")
    gStr.Add s, "fb.mailCopied"
    s = ""
    s = s & U("4F5C800590AE7BB1FF1A007B0030007D")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("4F6053EF4EE557284EFB4F5590AE7BB1FF087F51987590AE7BB1300100510051002090AE7BB13001004F00750074006C006F006F006B00207B49FF094E2D65B05EFA90AE4EF6540E53D190013002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("898173B0572862535F00753581114E0A76849ED88BA490AE4EF67A0B5E8F5417FF1F")
    s = s & vbCrLf
    s = s & U("FF085982679C6CA167098BBE7F6E90AE4EF67A0B5E8FFF0C53EF4EE5900962E9201C5426201DFF0C76F463A575287F51987590AE7BB153D19001FF09")
    gStr.Add s, "fb.mailPlain"
    gStr.Add U("90AE4EF653CD9988"), "fb.mailTitle"
    s = ""
    s = s & U("65E06CD562535F006D4F89C85668FF0C8BF7624B52A88BBF95EEFF1A")
    s = s & vbCrLf
    s = s & "{0}"
    gStr.Add s, "fb.noBrowser"
    gStr.Add U("65E06CD562535F0090AE4EF67A0B5E8FFF0C8BF776F463A553D190AE4EF65230FF1A007B0030007D"), "fb.noMailApp"
    gStr.Add U("77E253700020005600650063005300740061006D00700020007B0030007D002053CD9988"), "fb.subject"
    gStr.Add U("53CD9988"), "fb.title"
    gStr.Add U("7B2C007B0030007D9875"), "file.page"
    gStr.Add U("8BBF8FBE"), "fm.mac"
    gStr.Add U("8D446E907BA174065668"), "fm.win"
    gStr.Add U("0045004D00460020589E5F3A578B56FE514365874EF6"), "fmt.emf"
    gStr.Add U("0045004D005A0020538B7F297684589E5F3A578B56FE514365874EF6"), "fmt.emz"
    gStr.Add U("005000440046002077E291CF65876863"), "fmt.pdf"
    gStr.Add U("0050004E004700209AD852068FA873874F4D56FE"), "fmt.png"
    gStr.Add U("005300560047002053EF7F29653E77E291CF56FE5F62"), "fmt.svg"
    gStr.Add U("005600530044005800200056006900730069006F00207ED856FE"), "fmt.vsdx"
    gStr.Add U("0057004D0046002000570069006E0064006F00770073002056FE514365874EF6"), "fmt.wmf"
    s = ""
    s = s & U("77E253700020005600650063005300740061006D007000204F7F75288BF4660E")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("2460002090094E2D4E004E2A6216591A4E2A56FE5F62FF0C70B9201C5BFC51FA201DFF0C621676F463A570B95404683C5F0F630994AE30026CA1670990094E2D56FE5F6265F6FF0C53EF4EE55BFC51FA657498753002")
    s = s & vbCrLf
    s = s & U("24610020901098755BFC51FAFF1A6BCF98755E7B706F72474E004E2A65874EF6FF1B57285DE64FA77F29756556FE4E2D591A900998759762521953EA5BFC51FA6240900998753002")
    s = s & vbCrLf
    s = s & U("246200200045004D0046002056FE7247526F672CFF1A628A6240900956FE5F62590D5236621077E291CF56FE7247653E572865C18FB9FF0C539F56FE5F624E0D53D83002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("00B700200045004D00460020002F00200045004D005A0020002F00200057004D00460020002F0020005300560047FF1A77E291CF56FE30020050006F0077006500720050006F0069006E007400200032003000310033002053CA4EE54E0A90FD80FD5BFC51FA0020005300560047FF08003200330030003200204E4B524D76847248672C7ECF0020005800500053002051857F6E8F6C6362FF093002")
    s = s & vbCrLf
    s = s & U("00B70020005000440046FF1A9875976259275C0F4E0E56FE5F624E0081F4768477E291CF0020005000440046FF0C900254080020004C00610054006500583002")
    s = s & vbCrLf
    s = s & U("00B700200056005300440058FF1A0056006900730069006F002065874EF6FF0C56FE5F624E3A00200056006900730069006F0020539F751F5F6272B6300165875B5753EF7F168F91FF0C65E097005B8988C500200056006900730069006F3002")
    s = s & vbCrLf
    s = s & U("00B700200050004E0047FF1A0031003500300020002F00200033003000300020002F002000360030003000200044005000490020621681EA5B9A4E49FF0867009AD80020007B0030007D0020004400500049FF093002")
    s = s & vbCrLf
    s = s & U("00B7002080CC666FFF1A5728201C80CC666F201D83DC53554E2D900962E9900F660E3001767D827230017C73827230017F51683C621681EA5B9A4E49989C8272FF084E3B989882720020002F0020680751C682720020002F002051764ED6989C82720020002F002053D6827256680020002F0020005200470042FF09FF0C53EF8BBE900F660E5EA6548C8FB98DDD3002")
    s = s & vbCrLf
    s = s & U("00B700205BFC51FA65F64F1A663E793A8FDB5EA67A9753E3FF085F53524D6B659AA4548C767E52066BD4FF09FF0C63090020004500730063002053EF4EE553D66D883002")
    s = s & vbCrLf
    s = s & U("00B70020754C97628BED8A00FF1A5728201C51734E8E201D7EC476848BED8A0083DC53554E2D520763627B804F534E2D658730017E419AD44E2D658730010045006E0067006C006900730068300165E5672C8A9EFF0C7ACB5373751F65483002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("54085E765BFC51FA300152A080CC666F30010050004400460020548C901098755BFC51FA65F64F1A4E3465F6501F7528526A8D34677FFF0C5E7657286587672B4E3465F663D251654E009875FF085B8C6210540E81EA52A852209664FF093002")
    gStr.Add s, "help.ppt"
    gStr.Add U("54085E764E3A4E004E2A65874EF6"), "mode.merge"
    gStr.Add U("6BCF4E2A56FE5F624E004E2A65874EF6"), "mode.split"
    gStr.Add U("786E5B9A628A624067095BFC51FA8BBE7F6E6062590D4E3A9ED88BA4503C5417FF1F"), "msg.askReset"
    s = ""
    s = s & U("89815BFC51FA54EA4E9B5E7B706F7247FF1F")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("662FFF1A53EA5BFC51FA6240900976840020007B0030007D00209875")
    s = s & vbCrLf
    s = s & U("5426FF1A5BFC51FA516890E80020007B0031007D00209875")
    gStr.Add s, "msg.askWhich"
    s = ""
    s = s & U("6CA1670990094E2D4EFB4F5556FE5F623002")
    s = s & vbCrLf
    s = s & U("662F54265BFC51FA5F53524D5E7B706F72474E0A7684516890E856FE5F62FF1F")
    gStr.Add s, "msg.askWholeSlide"
    gStr.Add U("80CC666F900F660E5EA68BF78F93516500200030FF5E00310030003000204E4B95F4768465705B57FF080025FF093002"), "msg.badAlpha"
    gStr.Add U("81EA5B9A4E4952068FA873878BF78F93516500200031FF5E007B0030007D00204E4B95F4768465746570FF08004400500049FF093002"), "msg.badDpi"
    gStr.Add U("7F51683C95F48DDD8BF78F93516500200032FF5E00320030003000204E4B95F4768465705B57FF0853554F4DFF1A78C5FF093002"), "msg.badGrid"
    gStr.Add U("8FB98DDD8BF78F93516500200030FF5E00350030003000204E4B95F4768465705B57FF0853554F4DFF1A78C5FF093002"), "msg.badMargin"
    gStr.Add U("901098755BFC51FA59318D25FF1A007B0030007D"), "msg.batchFailed"
    gStr.Add U("63090020007B0030007D002000440050004900205BFC51FAFF0C8FD95F2056FE7EA64E3A0020007B0031007D002000D70020007B0032007D002050CF7D20FF087EA60020007B0033007D0020767E4E0750CF7D20FF09FF0C53EF80FD5F886162621651855B584E0D8DB330024ECD71367EE77EED5417FF1F"), "msg.bigPng"
    gStr.Add U("5DF253D66D885BFC51FA3002"), "msg.cancelled"
    gStr.Add U("5DF253D66D885BFC51FA30025DF25B8C621076840020007B0030007D00204E2A65874EF65DF24FDD7559FF1A"), "msg.cancelledPartial"
    s = ""
    s = s & U("65E06CD562535F00989C8272900962E95668FF1A007B0030007D")
    s = s & vbCrLf
    s = s & U("4E5F53EF4EE576F463A55728201C005200470042201D68464E2D8F935165989C82723002")
    gStr.Add s, "msg.colorFailed"
    gStr.Add U("5BFC51FA59318D25FF1A007B0030007D"), "msg.exportFailed"
    gStr.Add U("5DF25BFC51FA0020007B0030007D00204E2A65874EF6FF1A"), "msg.exported"
    s = ""
    s = s & U("53D682725668970089815728666E901A89C656FE4E2D62535F004E005F205E7B706F72473002")
    s = s & vbCrLf
    s = s & U("8BF7520763625230666E901A89C656FE540E518D70B9201C53D682725668201DFF0C7136540E57285E7B706F72474E0A70B951FB898162FE53D67684989C82723002")
    gStr.Add s, "msg.eyedropperView"
    gStr.Add U("987976EE4E3B9875FF1A007B0030007D"), "msg.home"
    gStr.Add U("8FD86CA167095BFC51FA8FC765874EF630025BFC51FA4E006B21540EFF0C8FD991CC4F1A62535F0067008FD14F7F7528768465874EF659393002"), "msg.noFolderYet"
    gStr.Add U("5F53524D5E7B706F72474E0A6CA1670953EF5BFC51FA768456FE5F623002"), "msg.noShapesSlide"
    gStr.Add U("624090095E7B706F72474E0A6CA1670953EF5BFC51FA768456FE5F623002"), "msg.noShapesSlides"
    gStr.Add U("6F14793A65877A3F4E2D6CA167095E7B706F72473002"), "msg.noSlides"
    s = ""
    s = s & U("6CA180FD751F62100020005300560047FF1A8FD94E2A00200050006F0077006500720050006F0069006E007400204E0D80FD76F463A55BFC51FA0020005300560047FF0C51857F6E8F6C63624E5F6CA167096210529FFF08970089817CFB7EDF81EA5E267684002000570069006E0064006F0077007300200050006F007700650072005300680065006C006CFF09FF0C753581114E0A4E5F6CA1670953EF75284E8E8F6C6362768400200049006E006B00730063006100700065002062160020004C0069006200720065004F006600660069006300653002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("662F542665394E3A5BFC51FA00200045004D0046FF1FFF084E5F53EF4EE55BFC51FA00200050004400460020540E752851764ED68F6F4EF68F6C6362FF09")
    gStr.Add s, "msg.noSvg"
    gStr.Add U("8BF75728666E901A89C656FE4E2D90094E2D89815BFC51FA768456FE5F623002"), "msg.normalView"
    gStr.Add U("8BF7514862535F006F14793A65877A3FFF0C5E7690094E2D89815BFC51FA768456FE5F623002"), "msg.openAndSelect"
    gStr.Add U("8BF7514862535F006F14793A65877A3F3002"), "msg.openPres"
    gStr.Add U("751F621000200045004D0046002056FE7247526F672C59318D25FF1A007B0030007D"), "msg.pictureFailed"
End Sub

Private Sub LoadStrings_zh_CN_3()
    Dim s As String
    gStr.Add U("8BF7514890094E2D56FE5F623002"), "msg.selectFirst"
    gStr.Add U("5DF28DF38FC70020007B0030007D00205F207A7A767D5E7B706F72473002"), "msg.skipped"
    s = ""
    s = s & U("51857F6E7684002000530056004700208F6C63628D858FC70020007B0030007D002079D24ECD672A5B8C6210FF0C5DF2505C6B623002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("56FE5F62975E5E38590D6742FF08592791CF65875B57300156FE72476216828270B9FF0965F653EF80FD51FA73B08FD979CD60C551B53002")
    s = s & vbCrLf
    s = s & U("662F54266539752800200049006E006B007300630061007000650020002F0020004C0069006200720065004F0066006600690063006500208F6C6362FF085982679C5DF25B8988C5FF0962165BFC51FA00200045004D0046FF1F")
    gStr.Add s, "msg.svgTimeout"
    gStr.Add U("6BCF6B21900962E94FDD5B584F4D7F6E"), "naming.ask"
    gStr.Add U("81EA52A84FDD5B5852309ED88BA465874EF65939"), "naming.auto"
    gStr.Add U("901098755BFC51FAFF08007B0030007D00209875FF0C007B0031007DFF09"), "prog.batch"
    gStr.Add U("FF085DF275280020007B0030007D002079D2FF09"), "prog.elapsed"
    gStr.Add U("63090020004500730063002053EF53D66D88"), "prog.escHint"
    gStr.Add U("6B6357285BFC51FA0020007B0030007D"), "prog.export1"
    gStr.Add U("6B6357285BFC51FA0020007B0030007D00204E2A0020007B0031007D002065874EF6"), "prog.exportN"
    gStr.Add U("53D682725668"), "prog.eyedropper"
    gStr.Add U("7B2C0020007B0030007D0020002F0020007B0031007D00204E2AFF1A007B0032007D"), "prog.item"
    gStr.Add U("7B2C0020007B0030007D0020002F0020007B0031007D00209875FF085E7B706F72470020007B0032007DFF09"), "prog.page"
    gStr.Add U("51C659074E2D2026"), "prog.preparing"
    gStr.Add U("7C738272"), "rb.vsBgBeige"
    gStr.Add U("81EA5B9A4E49989C82722026"), "rb.vsBgCustomMac"
    gStr.Add U("7F51683CFF08575068077EB8FF09"), "rb.vsBgGrid"
    gStr.Add U("900F660E"), "rb.vsBgNone"
    gStr.Add U("81EA5B9A4E49989C8272"), "rb.vsBgSepCustom"
    gStr.Add U("767D8272"), "rb.vsBgWhite"
    gStr.Add U("51734E8E77E25370"), "rb.vsBtnAbout"
    gStr.Add U("7248672C30014F5C80053001004D0049005400205F006E90534F8BAEFF1B77E253706C384E455F006E9030016C384E45514D8D393002"), "rb.vsBtnAbout.sup"
    gStr.Add U("51734E8E002077E253700020005600650063005300740061006D0070"), "rb.vsBtnAbout.tip"
    gStr.Add U("901098755BFC51FA"), "rb.vsBtnBatch"
    gStr.Add U("6BCF5F205E7B706F72475BFC51FA4E3A4E004E2A65874EF6FF084F7F7528201C9ED88BA4683C5F0F201DFF09300257285DE64FA77F29756556FE4E2D591A90099875976265F653EA5BFC51FA624090099875FF1B7A7A767D9875548C7A7A76846807989853604F4D7B264F1A88AB8DF38FC73002"), "rb.vsBtnBatch.sup"
    gStr.Add U("51734E8E8D5E8D4F2026"), "rb.vsBtnDonateInfo"
    gStr.Add U("589E5F3A578B56FE514365874EF6FF0C0057006F007200640020002F00200057005000530020002F00200050006F0077006500720050006F0069006E00740020901A7528FF0C8BBA658763D256FE999690093002"), "rb.vsBtnEMF.sup"
    gStr.Add U("5BFC51FA4E3A00200045004D0046"), "rb.vsBtnEMF.tip"
    gStr.Add U("0067007A006900700020538B7F29768400200045004D0046FF0C51855BB94E0E00200045004D0046002076F8540CFF0C65874EF666F45C0F3002"), "rb.vsBtnEMZ.sup"
    gStr.Add U("5BFC51FA4E3A00200045004D005A"), "rb.vsBtnEMZ.tip"
    gStr.Add U("5BFC51FA0020007B0030007D"), "rb.vsBtnExport"
    gStr.Add U("5BFC51FA90094E2D56FE5F62"), "rb.vsBtnExport.tip"
    gStr.Add U("53D682725668002800450029"), "rb.vsBtnEyedropper"
    gStr.Add U("53CD9988"), "rb.vsBtnFeedback"
    gStr.Add U("900962E953D1900190AE4EF67ED94F5C8005FF0C621657280020004700690074004800750062002063D04EA40020004900730073007500653002"), "rb.vsBtnFeedback.sup"
    gStr.Add U("53CD998895EE989862165EFA8BAE"), "rb.vsBtnFeedback.tip"
    gStr.Add U("9ED88BA465874EF659392026"), "rb.vsBtnFolder"
    s = ""
    s = s & U("201C81EA52A84FDD5B58201D548C201C901098755BFC51FA201D9ED88BA44F7F7528768465874EF65939FF1A")
    s = s & vbCrLf
    s = s & "{0}"
    gStr.Add s, "rb.vsBtnFolder.sup"
    gStr.Add U("4F7F75288BF4660E"), "rb.vsBtnHelp"
    gStr.Add U("5404630994AE548C8BBE7F6E76847B8089818BF4660E3002"), "rb.vsBtnHelp.sup"
    gStr.Add U("987976EE4E3B9875"), "rb.vsBtnHome"
    gStr.Add U("57286D4F89C856684E2D62535F0077E2537076840020004700690074004800750062002098759762FF1A4E0B8F7D65B07248672C300167E5770B6E907801548C66F465B065E55FD73002"), "rb.vsBtnHome.sup"
    gStr.Add U("663E793A5F53524D624067095BFC51FA8BBE7F6E3002"), "rb.vsBtnLauncher.sup"
    gStr.Add U("67E5770B5F53524D8BBE7F6E"), "rb.vsBtnLauncher.tip"
    gStr.Add U("51764ED6989C82720028004D00292026"), "rb.vsBtnMoreColors"
    gStr.Add U("62535F005BFC51FA65874EF65939"), "rb.vsBtnOpenFolder"
    gStr.Add U("5728007B0030007D4E2D62535F0067008FD14E006B215BFC51FA62405728768465874EF659393002"), "rb.vsBtnOpenFolder.sup"
    gStr.Add U("9875976259275C0F4E0E56FE5F624E0081F4768477E291CF00200050004400463002"), "rb.vsBtnPDF.sup"
    gStr.Add U("5BFC51FA4E3A0020005000440046"), "rb.vsBtnPDF.tip"
    gStr.Add U("9AD852068FA873874F4D56FEFF1B52068FA873875728201C5BFC51FA8BBE7F6E0020219200200050004E0047002052068FA87387201D4E2D900962E93002"), "rb.vsBtnPNG.sup"
End Sub

Private Sub LoadStrings_zh_CN_4()
    Dim s As String
    gStr.Add U("5BFC51FA4E3A00200050004E0047"), "rb.vsBtnPNG.tip"
    gStr.Add U("0045004D0046002056FE7247526F672C"), "rb.vsBtnPicture"
    gStr.Add U("628A6240900956FE5F62590D523662104E005F2077E291CF56FE7247653E572865C18FB9FF0C75286765201C51BB7ED3201D63927248FF1B539F56FE5F624FDD63014E0D53D83002"), "rb.vsBtnPicture.sup"
    gStr.Add U("6062590D9ED88BA48BBE7F6E"), "rb.vsBtnReset"
    gStr.Add U("628A624067095BFC51FA548C80CC666F8BBE7F6E6062590D4E3A9ED88BA4503CFF08754C97628BED8A004FDD63014E0D53D8FF093002"), "rb.vsBtnReset.sup"
    gStr.Add U("53EF7F29653E77E291CF56FE5F62FF0C6D4F89C8566853EF76F463A562535F0030020050006F0077006500720050006F0069006E007400200032003000310033002053CA4EE54E0A90FD80FD5BFC51FAFF1A65B07248672C76F463A55BFC51FAFF0C65E77248672C81EA52A87ECF0020005800500053002051857F6E8F6C6362FF0865875B574ECD662F65875B57FF09FF0C5E7657285BFC51FA540E68C067E565874EF63002"), "rb.vsBtnSVG.sup"
    gStr.Add U("5BFC51FA4E3A0020005300560047"), "rb.vsBtnSVG.tip"
    gStr.Add U("0056006900730069006F00207ED856FE65874EF6FF1A6BCF4E2A56FE5F6290FD8F6C63624E3A00200056006900730069006F0020539F751F5F6272B6FF085F6272B63001989C827230017EBF676130017BAD593430017EC45408FF09FF0C65875B5753EF76F463A57F168F91FF0C65E097005B8988C500200056006900730069006F300256FE7247300156FE88687B4965E06CD58F6C636276845BF98C614EE577E291CF56FE72475D4C51653002"), "rb.vsBtnVSDX.sup"
    gStr.Add U("5BFC51FA4E3A00200056005300440058"), "rb.vsBtnVSDX.tip"
    gStr.Add U("5B8C6210540E62535F0065874EF65939"), "rb.vsChkOpen"
    gStr.Add U("5BFC51FA5B8C6210540E57288D446E907BA1740656680020002F00208BBF8FBE4E2D5B9A4F4D52305BFC51FA768465874EF63002"), "rb.vsChkOpen.sup"
    gStr.Add U("00560053004400580020539F751F5F6272B6"), "rb.vsChkVsdx"
    s = ""
    s = s & U("52FE9009FF1A005600530044005800204E2D768456FE5F62662F00200056006900730069006F0020539F751F5F6272B6FF0C65875B5753EF76F463A57F168F91FF0865E097005B8988C500200056006900730069006FFF093002")
    s = s & vbCrLf
    s = s & U("4E0D52FE9009FF1A65744E2A56FE5F624F5C4E3A4E005F2000200045004D0046002077E291CF56FE72475D4C516500200056005300440058FF0C591689C24E0E00200050006F0077006500720050006F0069006E007400205B8C51684E0081F43002")
    gStr.Add s, "rb.vsChkVsdx.sup"
    gStr.Add U("0050004E0047002052068FA87387"), "rb.vsDdDpi"
    gStr.Add U("53EA5F7154CD00200050004E004700205BFC51FA3002671F520A53705237901A5E3889816C420020003300300030002000440050004900204EE54E0AFF0C7EBF676156FE53EF900900200036003000300020004400500049FF1B4E5F53EF4EE581EA5B9A4E49FF0867009AD8002000330030003000300020004400500049FF093002"), "rb.vsDdDpi.sup"
    gStr.Add U("9ED88BA4683C5F0F"), "rb.vsDdFormat"
    gStr.Add U("201C5BFC51FA201D548C201C901098755BFC51FA201D4F7F75287684683C5F0FFF084FDD5B585BF98BDD68464E2D4E5F53EF4EE54E3465F666F46539FF093002"), "rb.vsDdFormat.sup"
    gStr.Add U("591A4E2A56FE5F62"), "rb.vsDdMode"
    gStr.Add U("54085E764E3A4E004E2A65874EF6FF0C62166BCF4E2A56FE5F625206522B5BFC51FAFF0865874EF6540D53D681EA201C900962E97A97683C201D4E2D768456FE5F62540D79F0FF093002"), "rb.vsDdMode.sup"
    gStr.Add U("90094E2D591A4E2A56FE5F6265F6"), "rb.vsDdMode.tip"
    gStr.Add U("4FDD5B5865B95F0F"), "rb.vsDdNaming"
    gStr.Add U("6BCF6B215F3951FA5BF98BDD6846900962E94F4D7F6EFF0C621676F463A54FDD5B5852309ED88BA465874EF65939FF0865874EF6540D81EA52A85E2665F695F4FF093002"), "rb.vsDdNaming.sup"
    gStr.Add U("900F660E5EA60025"), "rb.vsEbAlpha"
    gStr.Add U("003000208868793A4E0D900F660EFF0C00310030003000208868793A5B8C5168900F660EFF1B53EA5F7154CD80CC666FFF0C4E0D5F7154CD56FE5F62672C8EAB3002"), "rb.vsEbAlpha.sup"
    gStr.Add U("80CC666F900F660E5EA6"), "rb.vsEbAlpha.tip"
    gStr.Add U("81EA5B9A4E490020004400500049"), "rb.vsEbDpi"
    gStr.Add U("0031FF5E003300300030003000204E4B95F47684657465703002201C0050004E0047002052068FA87387201D9009201C81EA5B9A4E49201D65F6751F6548300252068FA873878D8A9AD865874EF68D8A59273002"), "rb.vsEbDpi.sup"
    gStr.Add U("81EA5B9A4E4952068FA87387"), "rb.vsEbDpi.tip"
    gStr.Add U("7F51683C002878C50029"), "rb.vsEbGrid"
    gStr.Add U("7F51683C80CC666F7684683C5B5059275C0FFF0C53554F4D4E3A78C5FF1B6BCF002000350020683C4E0067616DF182727EBF3002"), "rb.vsEbGrid.sup"
    gStr.Add U("7F51683C95F48DDD"), "rb.vsEbGrid.tip"
    gStr.Add U("8FB98DDD002878C50029"), "rb.vsEbMargin"
    gStr.Add U("56FE5F6256DB54687559767D76845BBD5EA6FF0C53554F4D4E3A78C5FF080031002078C50020224800200030002E0033003500206BEB7C73FF093002003000208868793A7D278D3456FE5F623002"), "rb.vsEbMargin.sup"
    gStr.Add U("8FB98DDD"), "rb.vsEbMargin.tip"
    gStr.Add U("81EA5B9A4E4980CC666F8272"), "rb.vsEbRgb.tip"
    gStr.Add U("81EA5B9A4E49989C8272"), "rb.vsGalColor"
    s = ""
    s = s & U("5F53524D81EA5B9A4E4980CC666F8272FF1A007B0030007D")
    s = s & vbCrLf
    s = s & U("4ECE4E3B9898989C82723001680751C68272300167008FD14F7F75287684989C82724E2D900962E9FF0C62167528201C51764ED6989C8272201D201C53D682725668201DFF1B4E5F53EF4EE557280020005200470042002068464E2D8F93516500200052002C0047002C0042FF0859820020003200350035002C003200340038002C003200330035FF0962165341516D8FDB5236FF08598200200023004600460046003800450042FF09300290098272540E80CC666F81EA52A8520763624E3A201C81EA5B9A4E49989C8272201D3002")
    gStr.Add s, "rb.vsGalColor.sup"
    gStr.Add U("8D5E8D4F4F5C8005"), "rb.vsGalDonate"
    gStr.Add U("77E253706C384E45514D8D3930015F006E9030025982679C5B835E2E52304E864F60FF0C53EF4EE57528652F4ED85B9D62165FAE4FE1626B78018BF74F5C8005559D676F54965561FF085B8C516881EA613FFF0C4E0D5F7154CD4EFB4F55529F80FDFF093002"), "rb.vsGalDonate.sup"
    gStr.Add U("51734E8E"), "rb.vsGrpAbout"
    gStr.Add U("80CC666F"), "rb.vsGrpBackground"
    gStr.Add U("5BFC51FA90094E2D56FE5F62"), "rb.vsGrpExport"
    gStr.Add U("5BFC51FA8BBE7F6E"), "rb.vsGrpSettings"
    gStr.Add U("627991CF4E0E5DE55177"), "rb.vsGrpTools"
    gStr.Add U("80CC666FFF1A007B0030007D"), "rb.vsMnuBg"
    s = ""
    s = s & U("5F53524D80CC666FFF1A007B0030007D")
    s = s & vbCrLf
    s = s & U("5BFC51FA65F657AB572856FE5F624E0B9762768480CC666FFF1A900F660E3001767D827230017C73827230017F51683CFF08575068077EB8FF09621681EA5B9A4E49989C82723002201C81EA5B9A4E49989C8272201D4E2D53EF4ECE4E3B9898989C82723001680751C68272300167008FD14F7F75287684989C82724E2D900962E9FF0C62167528201C51764ED6989C8272201D201C53D682725668201D300270B951FB989C827265B9575753EF76F463A562535F00989C82729762677F3002")
    gStr.Add s, "rb.vsMnuBg.sup"
    gStr.Add U("80CC666F"), "rb.vsMnuBg.tip"
    gStr.Add U("0057006F007200640020002F00200057005000530020002F00200050006F0077006500720050006F0069006E00740020901A7528768477E291CF56FEFF0C8BBA658763D256FE999690093002"), "rb.vsMnuEMF.desc"
    gStr.Add U("0067007A006900700020538B7F29768400200045004D0046FF0C51855BB976F8540CFF0C4F5379EF66F45C0F3002"), "rb.vsMnuEMZ.desc"
    gStr.Add U("5F53524DFF1A007B0030007D30025207636277E253707684754C97628BED8A00FF087B804F534E2D658730017E419AD44E2D658730010045006E0067006C006900730068300165E5672C8A9EFF09FF0C7ACB5373751F65483002"), "rb.vsMnuLang.sup"
End Sub

Private Sub LoadStrings_zh_CN_5()
    Dim s As String
    gStr.Add U("754C97628BED8A000020002F0020004C0061006E00670075006100670065"), "rb.vsMnuLang.tip"
    gStr.Add U("9875976259275C0F4E0E56FE5F624E0081F4FF0C900254080020004C00610054006500580020548C62957A3F7CFB7EDF3002"), "rb.vsMnuPDF.desc"
    gStr.Add U("0031003500300020002F00200033003000300020002F002000360030003000200044005000490020621681EA5B9A4E49FF0867009AD8002000330030003000300020004400500049FF093002"), "rb.vsMnuPNG.desc"
    gStr.Add U("7F51987530010049006C006C007500730074007200610074006F007230010049006E006B007300630061007000650020901A7528FF1B0050006F0077006500720050006F0069006E007400200032003000310033002053CA4EE54E0A90FD80FD5BFC51FA3002"), "rb.vsMnuSVG.desc"
    gStr.Add U("65876863"), "rb.vsMnuSepDoc"
    gStr.Add U("4F4D56FE"), "rb.vsMnuSepRaster"
    gStr.Add U("77E291CF56FE"), "rb.vsMnuSepVector"
    gStr.Add U("0056006900730069006F0020539F751F5F6272B6FF0C65875B5753EF76F463A57F168F91FF1B65E097005B8988C500200056006900730069006F3002"), "rb.vsMnuVSDX.desc"
    gStr.Add U("517C5BB980018F6F4EF6FF1B4E0D652F6301900F660E548C6E1053D83002"), "rb.vsMnuWMF.desc"
    gStr.Add U("77E25370"), "rb.vsTab"
    gStr.Add U("5F53524D8BBE7F6E"), "set.current"
    gStr.Add U("FF08672A8BBE7F6EFF1A4F7F75286F14793A65877A3F6240572865874EF65939FF1B672A4FDD5B58621657284E917AEF65F64F7F7528684C9762FF09"), "set.folderNone"
    gStr.Add U("900962E9201C81EA52A84FDD5B58201D4F7F752876849ED88BA465874EF65939"), "set.pickDefaultFolder"
    gStr.Add U("5206679056FE5F627ED36784"), "step.analyze"
    gStr.Add U("57284E3465F65E7B706F72474E0A5408621056FE5F62300180CC666F4E0E8FB98DDD"), "step.compose"
    gStr.Add U("540862105B8C6210"), "step.composed"
    gStr.Add U("5B8C6210FF1A007B0030007D"), "step.doneFile"
    gStr.Add U("4F7F752800200049006E006B007300630061007000650020002F0020004C0069006200720065004F006600660069006300650020628A00200045004D004600208F6C63624E3A0020005300560047"), "step.emf2svg"
    gStr.Add U("538B7F294E3A00200045004D005AFF080067007A00690070FF09"), "step.emz"
    gStr.Add U("57285E7B706F72474E0A70B951FB898162FE53D67684989C8272FF0800360030002079D2518567096548FF09"), "step.eyedropper"
    gStr.Add U("68C067E55E764FEE590D0020005300560047"), "step.fixSvg"
    gStr.Add U("0049006E006B0073006300610070006500206B635728628A00200045004D004600208F6C63624E3A0020005300560047"), "step.inkscape"
    gStr.Add U("004C0069006200720065004F0066006600690063006500206B635728628A00200045004D004600208F6C63624E3A0020005300560047"), "step.libreoffice"
    gStr.Add U("521B5EFA4E0E56FE5F62540C683759275C0F76844E3465F665876863"), "step.pdfDoc"
    gStr.Add U("4EE577E291CF56FE5F625F0F7C988D34"), "step.pdfPaste"
    gStr.Add U("0050006F0077006500720050006F0069006E007400206B635728751F62100020005000440046"), "step.pdfWrite"
    gStr.Add U("6E3267D300200050004E0047FF1A007B0030007D002000D70020007B0031007D002050CF7D20FF08007B0032007D0020004400500049FF09"), "step.png"
    gStr.Add U("0050006F0077006500720050006F0069006E007400206B635728751F621000200045004D0046002077E291CF56FE"), "step.ppEmf"
    gStr.Add U("0050006F0077006500720050006F0069006E007400206B6357285BFC51FA0020005300560047"), "step.ppSvg"
    gStr.Add U("0050006F0077006500720050006F0069006E007400206B635728751F621000200057004D0046002056FE514365874EF6"), "step.ppWmf"
    gStr.Add U("51C6590756FE5F62FF08007B0030007D00204E2A5BF98C61FF09"), "step.prepare"
    gStr.Add U("751F621000200045004D0046002077E291CF56FEFF0856FE72476A215F0FFF09"), "step.vsdxPicture"
    gStr.Add U("751F621000200056006900730069006F00205F6272B60020007B0030007D0020002F0020007B0031007DFF1A007B0032007D"), "step.vsdxShape"
    gStr.Add U("519951FA00200056006900730069006F002065874EF6FF08007B0030007D00204E2A539F751F5F6272B6FF0C007B0031007D00205F2056FE7247FF09"), "step.vsdxWrite"
    gStr.Add U("519951FA00200056006900730069006F002065874EF6"), "step.vsdxWrite0"
    gStr.Add U("901A8FC700200058005000530020751F62100020005300560047FF1A007B0030007D"), "step.xps"
    gStr.Add U("51C659074E3465F65E7B706F7247FF08969085CF6BCD7248548C80CC666FFF09"), "step.xpsPrep"
    gStr.Add U("0050006F0077006500720050006F0069006E007400206B635728625353704E3A0020005800500053"), "step.xpsPrint"
    gStr.Add U("6253530500200056005300440058"), "step.zip"
    gStr.Add U("FF0C900F660E5EA60020007B0030007D0025"), "sum.alpha"
    gStr.Add U("80CC666FFF1A007B0030007D"), "sum.bg"
    gStr.Add U("0050004E0047002052068FA87387FF1A007B0030007D0020004400500049"), "sum.dpi"
    gStr.Add U("9ED88BA4683C5F0FFF1A007B0030007D"), "sum.format"
    gStr.Add U("63095F53524D8BBE7F6E5BFC51FA90094E2D56FE5F62"), "sum.head"
    gStr.Add U("754C97628BED8A00FF1A007B0030007D"), "sum.lang"
    gStr.Add U("8FB98DDDFF1A007B0030007D002078C5"), "sum.margin"
    gStr.Add U("591A4E2A56FE5F62FF1A007B0030007D"), "sum.mode"
    gStr.Add U("4FDD5B5865B95F0FFF1A007B0030007D"), "sum.naming"
    gStr.Add U("0056005300440058FF1A0056006900730069006F0020539F751F5F6272B6FF0865875B5753EF7F168F91FF09"), "sum.vsdxNative"
    gStr.Add U("0056005300440058FF1A5D4C516500200045004D0046002077E291CF56FE7247"), "sum.vsdxPicture"
End Sub

Private Sub LoadStrings_zh_CN_6()
    Dim s As String
    gStr.Add U("8F6C63625B8C6210"), "xps.done"
    gStr.Add U("8F6C636256FE5F6251437D200020007B0030007D0020002F0020007B0031007D"), "xps.elem"
    gStr.Add U("5D4C51655B574F53FF087B2C0020007B0030007D00204E2AFF09"), "xps.font"
    gStr.Add U("8BFB53D60020005800500053002065874EF6"), "xps.open"
    gStr.Add U("5206679098759762FF1A51710020007B0030007D00204E2A56FE5F6251437D20"), "xps.page"
    gStr.Add U("542F52A88F6C63627A0B5E8F"), "xps.start"
    gStr.Add U("519951FA0020005300560047002065874EF6"), "xps.write"
End Sub

Private Sub LoadStrings_zh_TW_1()
    Dim s As String
    gStr.Add U("4ECB97628A9E8A00FF1A007B0030007DFF0853EF5728529F80FD5340300C95DC65BC300D7D4476848A9E8A00907855AE4E2D520763DBFF09"), "about.lang"
    s = ""
    s = s & U("77E253700020005600650063005300740061006D0070002000200076007B0030007D")
    s = s & vbCrLf
    s = s & U("907853D653735370002020142014002057165F624E009375532F51FA00200045004D00460020002F00200053005600470020002F00200050004400460020002F002000560053004400580020002F00200050004E004700207B49683C5F0F")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("4F5C8005FF1A007B0031007D")
    s = s & vbCrLf
    s = s & U("4FE17BB1FF1A007B0032007D")
    s = s & vbCrLf
    s = s & U("99969801FF1A007B0033007D")
    s = s & vbCrLf
    s = s & U("958B6E9063886B0AFF1A004D004900540020004C006900630065006E00730065")
    s = s & vbCrLf
    s = s & "Copyright (c) {4} {1}"
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("77E25370627F8AFE6C384E45958B6E9030016C384E45514D8CBBFF1A6C92670965368CBB529F80FDFF0C6C9267095EE3544AFF0C4E0D806F7DB2653696C64EFB4F558CC78A0AFF0C516890E8539F59CB78BC572800200047006900740048007500620020516C958B3002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("5982679C5B8370BA4F607BC077014E8666429593FF0C6B618FCE9EDE9078529F80FD5340300C95DC65BC300D7D447684300C8D0A52A94F5C8005300DFF0C75285FAE4FE16216652F4ED85BF6638378BC8ACB4F5C8005559D676F5496556130028D0A52A95B8C516881EA9858FF0C4E0D5F7197FF4EFB4F55529F80FDFF1B4F607684652F63F46703752865BC63017E8C7DAD8B77548C6539903277E253703002")
    gStr.Add s, "about.text"
    gStr.Add U("95DC65BC0020007B0030007D"), "about.title"
    gStr.Add U("77E253700020005600650063005300740061006D0070"), "app.title"
    gStr.Add U("7C738272"), "bg.beige"
    gStr.Add U("81EA8A02984F8272"), "bg.custom"
    gStr.Add U("81EA8A02984F82720020007B0030007D"), "bg.customValue"
    gStr.Add U("683C7DDA"), "bg.grid"
    gStr.Add U("683C7DDAFF0895938DDD0020007B0030007D002078C5FF09"), "bg.gridValue"
    gStr.Add U("900F660E"), "bg.none"
    gStr.Add U("767D8272"), "bg.white"
    gStr.Add U("53D66D88"), "btn.cancel"
    gStr.Add U("78BA5B9A"), "btn.ok"
    gStr.Add U("984F8272683C5F0F793A4F8BFF1A003200350035002C003200340038002C0032003300350020621600200023004600460046003800450042"), "col.bad"
    gStr.Add U("6DF182720020007B0030007D0025"), "col.darker"
    gStr.Add U("63094E004E0B8272584A90788272FF0C630951694E0B76F463A578BA5B9AFF1B4E5F53EF8F38516500200052002C0047002C004200206216002000230052005200470047004200423002"), "col.hint"
    gStr.Add U("6DE182720020007B0030007D0025"), "col.lighter"
    gStr.Add U("8F3851657DE8865F907864C7984F8272FF0C62168F38516500200052002C0047002C00420020002F00200023005200520047004700420042FF1A"), "col.listPrompt"
    gStr.Add U("67008FD14F7F75287684984F8272"), "col.recent"
    gStr.Add U("67008FD14F7F75287684984F8272FF1A007B0030007D"), "col.recentItem"
    gStr.Add U("6A196E968272"), "col.standard"
    gStr.Add U("6A196E968272FF1A007B0030007D"), "col.stdItem"
    gStr.Add U("6DF17D05007C7D058272007C6A598272007C9EC38272007C6DFA7DA0007C7DA08272007C6DFA85CD007C85CD8272007C6DF185CD007C7D2B8272"), "col.stdNames"
    gStr.Add U("7CFB7D71984F82722026"), "col.system"
    gStr.Add U("7CFB7D71984F8272"), "col.systemPicked"
    gStr.Add U("4E3B984C984F8272"), "col.theme"
    gStr.Add U("767D8272FF0C80CC666F00200031007C9ED18272FF0C65875B5700200031007C80CC666F00200032007C65875B5700200032007C8F14827200200031007C8F14827200200032007C8F14827200200033007C8F14827200200034007C8F14827200200035007C8F14827200200036"), "col.themeNames"
    gStr.Add U("007B0030007DFF0C007B0031007D"), "col.tint"
    gStr.Add U("81EA8A02984F8272"), "col.title"
    gStr.Add U("984F8272503C"), "col.value"
    gStr.Add U("007B0030007DFF08007B0031007DFF09"), "col.withValue"
    gStr.Add U("6A94540DFF084E0D542B526F6A94540DFF09FF1A"), "dlg.fileName"
    gStr.Add U("907864C7532F51FA8CC76599593E"), "dlg.pickFolder"
    gStr.Add U("907864C790109801532F51FA76848CC76599593EFF08683C5F0FFF1A007B0030007DFF09"), "dlg.pickFolderBatch"
    gStr.Add U("907864C7532F51FA8CC76599593EFF086BCF500B57165F624E00500B6A946848FF09"), "dlg.pickFolderEach"
    gStr.Add U("532F51FA70BA0020007B0030007D"), "dlg.saveAs"
    gStr.Add U("532F51FA907853D657165F62"), "dlg.saveTitle"
    gStr.Add U("652F4ED85BF6"), "don.alipay"
    s = ""
    s = s & U("8B1D8B1D4F609858610F652F63F477E25370FF01")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("8ACB7528007B0030007D638363CF300C8D0A52A94F5C8005300D4E2D5C0D61C976840051005200200043006F00640065FF0C91D1984D96A8610F3002")
    s = s & vbCrLf
    s = s & U("8D0A52A95B8C516881EA9858FF0C77E25370768462406709529F80FD90FD6C384E45514D8CBB3001958B6E903002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("5982679C65B94FBFFF0C4E5F6B618FCE5728002000470069007400480075006200207D665C0868489EDE4E00500B00200053007400610072FF0C6216628A5B8363A885A67D66970089817684540C5B78548C540C4E8B3002")
    gStr.Add s, "don.thanks"
    gStr.Add U("8D0A52A9"), "don.title"
    gStr.Add U("5FAE4FE1"), "don.wechat"
    gStr.Add U("81EA8A02"), "dpi.custom"
    gStr.Add U("71216CD5900F904E526A8CBC7C3F890788FD6240907857165F62FF08526A8CBC7C3F53EF80FD88AB51764ED67A0B5F0F4F547528FF09FF0C8ACB91CD8A663002"), "err.clipboard"
    gStr.Add U("0045004D005A002058D37E2E593165573002"), "err.emz"
    gStr.Add U("6240907872694EF671216CD57D445408532F51FA3002"), "err.group"
    gStr.Add U("627E4E0D52300020006D00610063004F005300208F1452A963074EE478BC0020007B0030007DFF0C8ACB91CD65B057F7884C0020006D00610063004F005300205B8988DD56683002"), "err.macHelper"
    gStr.Add U("FF080050006F0077006500720050006F0069006E007400200066006F00720020004D00610063002097008981002000310036002E003800320020621666F49AD87248672C624D80FD532F51FA57165F62FF09"), "err.macVersion"
    gStr.Add U("71216CD5628A6A9468485BEB5165FF1A007B0030007D"), "err.macWrite"
    gStr.Add U("6C9267097522751F6A946848FF1A007B0030007D"), "err.noFile"
    gStr.Add U("0050004400460020532F51FA59316557FF1A007B0030007D"), "err.pdf"
End Sub

Private Sub LoadStrings_zh_TW_2()
    Dim s As String
    gStr.Add U("62535305002000560053004400580020593165573002"), "err.zip"
    gStr.Add U("6253530500200056005300440058002059316557FF08970089817CFB7D7181EA5E367684002000570069006E0064006F0077007300200050006F007700650072005300680065006C006CFF093002"), "err.zipPs"
    s = ""
    s = s & U("8ACB907864C7610F898B56DE994B65B95F0FFF1A")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("3010662F301150B3900190F54EF67D664F5C8005FF08007B0030007DFF09")
    s = s & vbCrLf
    s = s & U("30105426301157280020004700690074004800750062002063D04EA4002000490073007300750065FF0897008981002000470069007400480075006200205E33865FFF0C554F984C6703516C958B986F793AFF09")
    s = s & vbCrLf
    s = s & U("301053D66D8830118FD456DE")
    gStr.Add s, "fb.ask"
    s = ""
    s = s & U("002A002A554F984C63CF8FF00020002F00205EFA8B70FF1A002A002A")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("002A002A5FA973FE6B659A5FFF1A002A002A")
    s = s & vbCrLf
    s = s & "1. "
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("002A002A57F7884C74B05883FF1A002A002A")
    s = s & vbCrLf
    s = s & "```"
    s = s & vbCrLf
    s = s & "{0}"
    s = s & vbCrLf
    s = s & "```"
    s = s & vbCrLf
    gStr.Add s, "fb.issueBody"
    gStr.Add U("005B610F898B56DE994B005D0020"), "fb.issueTitle"
    s = ""
    s = s & U("4F5C80054FE17BB15DF2890788FD5230526A8CBC7C3FFF1A007B0030007D")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("4F6053EF4EE557284EFB4F554FE17BB1FF087DB298014FE17BB130010051005100204FE17BB13001004F00750074006C006F006F006B00207B49FF094E2D65B0589E90F54EF630018CBC4E0A65364EF64EBA5F8C50B390013002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("898173FE5728958B555F96FB81664E0A768498108A2D90F54EF67A0B5F0F55CEFF1F")
    s = s & vbCrLf
    s = s & U("FF085982679C6C9267098A2D5B9A90F54EF67A0B5F0FFF0C53EF4EE5907864C7300C5426300DFF0C76F463A575287DB298014FE17BB150B39001FF09")
    gStr.Add s, "fb.mailCopied"
    s = ""
    s = s & U("4F5C80054FE17BB1FF1A007B0030007D")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("4F6053EF4EE557284EFB4F554FE17BB1FF087DB298014FE17BB130010051005100204FE17BB13001004F00750074006C006F006F006B00207B49FF094E2D65B0589E90F54EF65F8C50B390013002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("898173FE5728958B555F96FB81664E0A768498108A2D90F54EF67A0B5F0F55CEFF1F")
    s = s & vbCrLf
    s = s & U("FF085982679C6C9267098A2D5B9A90F54EF67A0B5F0FFF0C53EF4EE5907864C7300C5426300DFF0C76F463A575287DB298014FE17BB150B39001FF09")
    gStr.Add s, "fb.mailPlain"
    gStr.Add U("90F54EF6610F898B56DE994B"), "fb.mailTitle"
    s = ""
    s = s & U("71216CD5958B555F700F89BD5668FF0C8ACB624B52D58A2A554FFF1A")
    s = s & vbCrLf
    s = s & "{0}"
    gStr.Add s, "fb.noBrowser"
    gStr.Add U("71216CD5958B555F90F54EF67A0B5F0FFF0C8ACB76F463A5767C90F54EF65230FF1A007B0030007D"), "fb.noMailApp"
    gStr.Add U("77E253700020005600650063005300740061006D00700020007B0030007D0020610F898B56DE994B"), "fb.subject"
    gStr.Add U("610F898B56DE994B"), "fb.title"
    gStr.Add U("7B2C007B0030007D9801"), "file.page"
    gStr.Add "Finder", "fm.mac"
    gStr.Add U("6A9468487E3D7BA1"), "fm.win"
    gStr.Add U("0045004D00460020589E5F37578B571651436A946848"), "fmt.emf"
    gStr.Add U("0045004D005A002058D37E2E7684589E5F37578B571651436A946848"), "fmt.emz"
    gStr.Add U("0050004400460020541191CF65874EF6"), "fmt.pdf"
    gStr.Add U("0050004E004700209AD889E367905EA69EDE96635716"), "fmt.png"
    gStr.Add U("005300560047002053EF7E2E653E541191CF57165F62"), "fmt.svg"
    gStr.Add U("005600530044005800200056006900730069006F00207E6A5716"), "fmt.vsdx"
    gStr.Add U("0057004D0046002000570069006E0064006F007700730020571651436A946848"), "fmt.wmf"
    s = ""
    s = s & U("77E253700020005600650063005300740061006D007000204F7F75288AAA660E")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("24600020907853D64E00500B6216591A500B57165F62FF0C9EDE300C532F51FA300DFF0C621676F463A59EDE5404683C5F0F6309921530026C926709907853D657165F626642FF0C53EF4EE5532F51FA657498013002")
    s = s & vbCrLf
    s = s & U("2461002090109801532F51FAFF1A6BCF980162955F7172474E00500B6A946848FF1B57285DE650747E2E57164E2D591A907898019762524753EA532F51FA6240907898013002")
    s = s & vbCrLf
    s = s & U("246200200045004D0046002057167247526F672CFF1A628A6240907857165F62890788FD6210541191CF57167247653E572865C1908AFF0C539F57165F624E0D8B8A3002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("00B700200045004D00460020002F00200045004D005A0020002F00200057004D00460020002F0020005300560047FF1A541191CF571630020050006F0077006500720050006F0069006E007400200032003000310033002053CA4EE54E0A90FD80FD532F51FA0020005300560047FF08003200330030003200204E4B524D76847248672C7D930020005800500053002051675EFA8F4963DBFF093002")
    s = s & vbCrLf
    s = s & U("00B70020005000440046FF1A9801976259275C0F820757165F624E0081F47684541191CF0020005000440046FF0C906954080020004C00610054006500583002")
    s = s & vbCrLf
    s = s & U("00B700200056005300440058FF1A0056006900730069006F00206A946848FF0C57165F6270BA00200056006900730069006F0020539F751F5F6272C0300165875B5753EF7DE88F2FFF0C712197005B8988DD00200056006900730069006F3002")
    s = s & vbCrLf
    s = s & U("00B700200050004E0047FF1A0031003500300020002F00200033003000300020002F002000360030003000200044005000490020621681EA8A02FF0867009AD80020007B0030007D0020004400500049FF093002")
    s = s & vbCrLf
    s = s & U("00B7002080CC666FFF1A5728300C80CC666F300D907855AE4E2D907864C7900F660E3001767D827230017C7382723001683C7DDA621681EA8A02984F8272FF084E3B984C82720020002F00206A196E9682720020002F002051764ED6984F82720020002F002053D6827256680020002F0020005200470042FF09FF0C53EF8A2D900F660E5EA6548C908A754C3002")
    s = s & vbCrLf
    s = s & U("00B70020532F51FA66426703986F793A90325EA689967A97FF0876EE524D6B659A5F548C767E52066BD4FF09FF0C63090020004500730063002053EF4EE553D66D883002")
    s = s & vbCrLf
    s = s & U("00B700204ECB97628A9E8A00FF1A5728300C95DC65BC300D7D4476848A9E8A00907855AE4E2D520763DB7C219AD44E2D658730017E419AD44E2D658730010045006E0067006C006900730068300165E5672C8A9EFF0C7ACB5373751F65483002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("54084F75532F51FA300152A080CC666F30010050004400460020548C90109801532F51FA6642670381E86642501F7528526A8CBC7C3FFF0C4E2657286587672B81E8664263D251654E009801FF085B8C62105F8C81EA52D5522A9664FF093002")
    gStr.Add s, "help.ppt"
    gStr.Add U("54084F7570BA4E00500B6A946848"), "mode.merge"
    gStr.Add U("6BCF500B57165F624E00500B6A946848"), "mode.split"
    gStr.Add U("78BA5B9A628A62406709532F51FA8A2D5B9A60625FA970BA98108A2D503C55CEFF1F"), "msg.askReset"
    s = ""
    s = s & U("8981532F51FA54EA4E9B62955F717247FF1F")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("662FFF1A53EA532F51FA6240907876840020007B0030007D00209801")
    s = s & vbCrLf
    s = s & U("5426FF1A532F51FA516890E80020007B0031007D00209801")
    gStr.Add s, "msg.askWhich"
    s = ""
    s = s & U("6C926709907853D64EFB4F5557165F623002")
    s = s & vbCrLf
    s = s & U("662F5426532F51FA76EE524D62955F7172474E0A7684516890E857165F62FF1F")
    gStr.Add s, "msg.askWholeSlide"
    gStr.Add U("80CC666F900F660E5EA68ACB8F38516500200030FF5E00310030003000204E4B9593768465785B57FF080025FF093002"), "msg.badAlpha"
    gStr.Add U("81EA8A0289E367905EA68ACB8F38516500200031FF5E007B0030007D00204E4B9593768465746578FF08004400500049FF093002"), "msg.badDpi"
    gStr.Add U("683C7DDA95938DDD8ACB8F38516500200032FF5E00320030003000204E4B9593768465785B57FF0855AE4F4DFF1A78C5FF093002"), "msg.badGrid"
    gStr.Add U("908A754C8ACB8F38516500200030FF5E00350030003000204E4B9593768465785B57FF0855AE4F4DFF1A78C5FF093002"), "msg.badMargin"
    gStr.Add U("90109801532F51FA59316557FF1A007B0030007D"), "msg.batchFailed"
    gStr.Add U("63090020007B0030007D00200044005000490020532F51FAFF0C90195F3557167D0470BA0020007B0031007D002000D70020007B0032007D0020756B7D20FF087D040020007B0033007D0020767E842C756B7D20FF09FF0C53EF80FD5F88616262168A1861B69AD44E0D8DB330024ECD71367E7C7E8C55CEFF1F"), "msg.bigPng"
    gStr.Add U("5DF253D66D88532F51FA3002"), "msg.cancelled"
    gStr.Add U("5DF253D66D88532F51FA30025DF25B8C621076840020007B0030007D0020500B6A9468485DF24FDD7559FF1A"), "msg.cancelledPartial"
    s = ""
    s = s & U("71216CD5958B555F984F8272907864C75668FF1A007B0030007D")
    s = s & vbCrLf
    s = s & U("4E5F53EF4EE576F463A55728300C005200470042300D68464E2D8F385165984F82723002")
    gStr.Add s, "msg.colorFailed"
    gStr.Add U("532F51FA59316557FF1A007B0030007D"), "msg.exportFailed"
    gStr.Add U("5DF2532F51FA0020007B0030007D0020500B6A946848FF1A"), "msg.exported"
    s = ""
    s = s & U("53D682725668970089815728666E901A6AA289964E2D958B555F4E005F3562955F7172473002")
    s = s & vbCrLf
    s = s & U("8ACB520763DB5230666E901A6AA289965F8C518D9EDE300C53D682725668300DFF0C71365F8C572862955F7172474E0A9EDE9078898162FE53D67684984F82723002")
    gStr.Add s, "msg.eyedropperView"
    gStr.Add U("5C08684899969801FF1A007B0030007D"), "msg.home"
    gStr.Add U("90846C926709532F51FA904E6A9468483002532F51FA4E006B215F8CFF0C901988E16703958B555F67008FD14F7F752876848CC76599593E3002"), "msg.noFolderYet"
    gStr.Add U("76EE524D62955F7172474E0A6C92670953EF532F51FA768457165F623002"), "msg.noShapesSlide"
    gStr.Add U("6240907862955F7172474E0A6C92670953EF532F51FA768457165F623002"), "msg.noShapesSlides"
    gStr.Add U("7C2158314E2D6C92670962955F7172473002"), "msg.noSlides"
    s = ""
    s = s & U("6C9280FD7522751F0020005300560047FF1A9019500B00200050006F0077006500720050006F0069006E007400204E0D80FD76F463A5532F51FA0020005300560047FF0C51675EFA8F4963DB4E5F6C9267096210529FFF08970089817CFB7D7181EA5E367684002000570069006E0064006F0077007300200050006F007700650072005300680065006C006CFF09FF0C96FB81664E0A4E5F6C92670953EF752865BC8F4963DB768400200049006E006B00730063006100700065002062160020004C0069006200720065004F006600660069006300653002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("662F5426653970BA532F51FA00200045004D0046FF1FFF084E5F53EF4EE5532F51FA002000500044004600205F8C752851764ED68EDF9AD48F4963DBFF09")
    gStr.Add s, "msg.noSvg"
    gStr.Add U("8ACB5728666E901A6AA289964E2D907853D68981532F51FA768457165F623002"), "msg.normalView"
    gStr.Add U("8ACB5148958B555F7C215831FF0C4E26907853D68981532F51FA768457165F623002"), "msg.openAndSelect"
    gStr.Add U("8ACB5148958B555F7C2158313002"), "msg.openPres"
    gStr.Add U("7522751F00200045004D0046002057167247526F672C59316557FF1A007B0030007D"), "msg.pictureFailed"
End Sub

Private Sub LoadStrings_zh_TW_3()
    Dim s As String
    gStr.Add U("8ACB5148907853D657165F623002"), "msg.selectFirst"
    gStr.Add U("5DF28DF3904E0020007B0030007D00205F357A7A767D62955F7172473002"), "msg.skipped"
    s = ""
    s = s & U("51675EFA7684002000530056004700208F4963DB8D85904E0020007B0030007D002079D24ECD672A5B8C6210FF0C5DF2505C6B623002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("57165F62975E5E38890796DCFF08592791CF65875B5730015716724762167BC09EDEFF09664253EF80FD51FA73FE90197A2E60C56CC13002")
    s = s & vbCrLf
    s = s & U("662F54266539752800200049006E006B007300630061007000650020002F0020004C0069006200720065004F0066006600690063006500208F4963DBFF085982679C5DF25B8988DDFF096216532F51FA00200045004D0046FF1F")
    gStr.Add s, "msg.svgTimeout"
    gStr.Add U("6BCF6B21907864C751325B584F4D7F6E"), "naming.ask"
    gStr.Add U("81EA52D551325B58523098108A2D8CC76599593E"), "naming.auto"
    gStr.Add U("90109801532F51FAFF08007B0030007D00209801FF0C007B0031007DFF09"), "prog.batch"
    gStr.Add U("FF085DF275280020007B0030007D002079D2FF09"), "prog.elapsed"
    gStr.Add U("63090020004500730063002053EF53D66D88"), "prog.escHint"
    gStr.Add U("6B635728532F51FA0020007B0030007D"), "prog.export1"
    gStr.Add U("6B635728532F51FA0020007B0030007D0020500B0020007B0031007D00206A946848"), "prog.exportN"
    gStr.Add U("53D682725668"), "prog.eyedropper"
    gStr.Add U("7B2C0020007B0030007D0020002F0020007B0031007D0020500BFF1A007B0032007D"), "prog.item"
    gStr.Add U("7B2C0020007B0030007D0020002F0020007B0031007D00209801FF0862955F7172470020007B0032007DFF09"), "prog.page"
    gStr.Add U("6E9650994E2D2026"), "prog.preparing"
    gStr.Add U("7C738272"), "rb.vsBgBeige"
    gStr.Add U("81EA8A02984F82722026"), "rb.vsBgCustomMac"
    gStr.Add U("683C7DDAFF085EA76A197D19FF09"), "rb.vsBgGrid"
    gStr.Add U("900F660E"), "rb.vsBgNone"
    gStr.Add U("81EA8A02984F8272"), "rb.vsBgSepCustom"
    gStr.Add U("767D8272"), "rb.vsBgWhite"
    gStr.Add U("95DC65BC77E25370"), "rb.vsBtnAbout"
    gStr.Add U("7248672C30014F5C80053001004D004900540020958B6E9063886B0AFF1B77E253706C384E45958B6E9030016C384E45514D8CBB3002"), "rb.vsBtnAbout.sup"
    gStr.Add U("95DC65BC002077E253700020005600650063005300740061006D0070"), "rb.vsBtnAbout.tip"
    gStr.Add U("90109801532F51FA"), "rb.vsBtnBatch"
    gStr.Add U("6BCF5F3562955F717247532F51FA70BA4E00500B6A946848FF084F7F7528300C98108A2D683C5F0F300DFF09300257285DE650747E2E57164E2D591A907898019762664253EA532F51FA624090789801FF1B7A7A767D9801548C7A7A76846A19984C981075594F4D7F6E670388AB8DF3904E3002"), "rb.vsBtnBatch.sup"
    gStr.Add U("95DC65BC8D0A52A92026"), "rb.vsBtnDonateInfo"
    gStr.Add U("589E5F37578B571651436A946848FF0C0057006F007200640020002F00200057005000530020002F00200050006F0077006500720050006F0069006E00740020901A7528FF0C8AD6658763D25716999690783002"), "rb.vsBtnEMF.sup"
    gStr.Add U("532F51FA70BA00200045004D0046"), "rb.vsBtnEMF.tip"
    gStr.Add U("0067007A00690070002058D37E2E768400200045004D0046FF0C51675BB9820700200045004D0046002076F8540CFF0C6A94684866F45C0F3002"), "rb.vsBtnEMZ.sup"
    gStr.Add U("532F51FA70BA00200045004D005A"), "rb.vsBtnEMZ.tip"
    gStr.Add U("532F51FA0020007B0030007D"), "rb.vsBtnExport"
    gStr.Add U("532F51FA907853D657165F62"), "rb.vsBtnExport.tip"
    gStr.Add U("53D682725668002800450029"), "rb.vsBtnEyedropper"
    gStr.Add U("610F898B56DE994B"), "rb.vsBtnFeedback"
    gStr.Add U("907864C750B3900190F54EF67D664F5C8005FF0C621657280020004700690074004800750062002063D04EA40020004900730073007500653002"), "rb.vsBtnFeedback.sup"
    gStr.Add U("610F898B56DE994B554F984C62165EFA8B70"), "rb.vsBtnFeedback.tip"
    gStr.Add U("98108A2D8CC76599593E2026"), "rb.vsBtnFolder"
    s = ""
    s = s & U("300C81EA52D551325B58300D548C300C90109801532F51FA300D98108A2D4F7F752876848CC76599593EFF1A")
    s = s & vbCrLf
    s = s & "{0}"
    gStr.Add s, "rb.vsBtnFolder.sup"
    gStr.Add U("4F7F75288AAA660E"), "rb.vsBtnHelp"
    gStr.Add U("540463099215548C8A2D5B9A76847C2189818AAA660E3002"), "rb.vsBtnHelp.sup"
    gStr.Add U("5C08684899969801"), "rb.vsBtnHome"
    gStr.Add U("5728700F89BD56684E2D958B555F77E2537076840020004700690074004800750062002098019762FF1A4E0B8F0965B07248672C30016AA28996539F59CB78BC548C66F465B065E58A8C3002"), "rb.vsBtnHome.sup"
    gStr.Add U("986F793A76EE524D62406709532F51FA8A2D5B9A3002"), "rb.vsBtnLauncher.sup"
    gStr.Add U("6AA2899676EE524D8A2D5B9A"), "rb.vsBtnLauncher.tip"
    gStr.Add U("51764ED6984F82720028004D00292026"), "rb.vsBtnMoreColors"
    gStr.Add U("958B555F532F51FA8CC76599593E"), "rb.vsBtnOpenFolder"
    gStr.Add U("5728007B0030007D4E2D958B555F67008FD14E006B21532F51FA6240572876848CC76599593E3002"), "rb.vsBtnOpenFolder.sup"
    gStr.Add U("9801976259275C0F820757165F624E0081F47684541191CF00200050004400463002"), "rb.vsBtnPDF.sup"
    gStr.Add U("532F51FA70BA0020005000440046"), "rb.vsBtnPDF.tip"
    gStr.Add U("9AD889E367905EA69EDE96635716FF1B89E367905EA65728300C532F51FA8A2D5B9A0020219200200050004E0047002089E367905EA6300D4E2D907864C73002"), "rb.vsBtnPNG.sup"
End Sub

Private Sub LoadStrings_zh_TW_4()
    Dim s As String
    gStr.Add U("532F51FA70BA00200050004E0047"), "rb.vsBtnPNG.tip"
    gStr.Add U("0045004D0046002057167247526F672C"), "rb.vsBtnPicture"
    gStr.Add U("628A6240907857165F62890788FD62104E005F35541191CF57167247653E572865C1908AFF0C75284F86300C51CD7D50300D63927248FF1B539F57165F624FDD63014E0D8B8A3002"), "rb.vsBtnPicture.sup"
    gStr.Add U("60625FA998108A2D8A2D5B9A"), "rb.vsBtnReset"
    gStr.Add U("628A62406709532F51FA548C80CC666F8A2D5B9A60625FA970BA98108A2D503CFF084ECB97628A9E8A004FDD63014E0D8B8AFF093002"), "rb.vsBtnReset.sup"
    gStr.Add U("53EF7E2E653E541191CF57165F62FF0C700F89BD566853EF76F463A5958B555F30020050006F0077006500720050006F0069006E007400200032003000310033002053CA4EE54E0A90FD80FD532F51FAFF1A65B07248672C76F463A5532F51FAFF0C820A7248672C81EA52D57D930020005800500053002051675EFA8F4963DBFF0865875B574ECD662F65875B57FF09FF0C4E265728532F51FA5F8C6AA267E56A9468483002"), "rb.vsBtnSVG.sup"
    gStr.Add U("532F51FA70BA0020005300560047"), "rb.vsBtnSVG.tip"
    gStr.Add U("0056006900730069006F00207E6A57166A946848FF1A6BCF500B57165F6290FD8F4963DB70BA00200056006900730069006F0020539F751F5F6272C0FF085F6272C03001984F827230017DDA689D30017BAD982D30017D445408FF09FF0C65875B5753EF76F463A57DE88F2FFF0C712197005B8988DD00200056006900730069006F3002571672473001571688687B4971216CD58F4963DB768472694EF64EE5541191CF571672475D4C51653002"), "rb.vsBtnVSDX.sup"
    gStr.Add U("532F51FA70BA00200056005300440058"), "rb.vsBtnVSDX.tip"
    gStr.Add U("5B8C62105F8C958B555F8CC76599593E"), "rb.vsChkOpen"
    gStr.Add U("532F51FA5B8C62105F8C57286A9468487E3D7BA10020002F002000460069006E00640065007200204E2D5B9A4F4D5230532F51FA76846A9468483002"), "rb.vsChkOpen.sup"
    gStr.Add U("00560053004400580020539F751F5F6272C0"), "rb.vsChkVsdx"
    s = ""
    s = s & U("52FE9078FF1A005600530044005800204E2D768457165F62662F00200056006900730069006F0020539F751F5F6272C0FF0C65875B5753EF76F463A57DE88F2FFF08712197005B8988DD00200056006900730069006FFF093002")
    s = s & vbCrLf
    s = s & U("4E0D52FE9078FF1A6574500B57165F624F5C70BA4E005F3500200045004D00460020541191CF571672475D4C516500200056005300440058FF0C591689C0820700200050006F0077006500720050006F0069006E007400205B8C51684E0081F43002")
    gStr.Add s, "rb.vsChkVsdx.sup"
    gStr.Add U("0050004E0047002089E367905EA6"), "rb.vsDdDpi"
    gStr.Add U("96BB5F7197FF00200050004E00470020532F51FA3002671F520A53705237901A5E3889816C420020003300300030002000440050004900204EE54E0AFF0C7DDA689D571653EF907800200036003000300020004400500049FF1B4E5F53EF4EE581EA8A02FF0867009AD8002000330030003000300020004400500049FF093002"), "rb.vsDdDpi.sup"
    gStr.Add U("98108A2D683C5F0F"), "rb.vsDdFormat"
    gStr.Add U("300C532F51FA300D548C300C90109801532F51FA300D4F7F75287684683C5F0FFF0851325B585C0D8A7165B9584A4E2D4E5F53EF4EE581E8664266F46539FF093002"), "rb.vsDdFormat.sup"
    gStr.Add U("591A500B57165F62"), "rb.vsDdMode"
    gStr.Add U("54084F7570BA4E00500B6A946848FF0C62166BCF500B57165F6252065225532F51FAFF086A94540D53D681EA300C907864C77A97683C300D4E2D768457165F62540D7A31FF093002"), "rb.vsDdMode.sup"
    gStr.Add U("907853D6591A500B57165F626642"), "rb.vsDdMode.tip"
    gStr.Add U("51325B5865B95F0F"), "rb.vsDdNaming"
    gStr.Add U("6BCF6B218DF351FA5C0D8A7165B9584A907864C74F4D7F6EFF0C621676F463A551325B58523098108A2D8CC76599593EFF086A94540D81EA52D55E3666429593FF093002"), "rb.vsDdNaming.sup"
    gStr.Add U("900F660E5EA60025"), "rb.vsEbAlpha"
    gStr.Add U("003000208868793A4E0D900F660EFF0C00310030003000208868793A5B8C5168900F660EFF1B96BB5F7197FF80CC666FFF0C4E0D5F7197FF57165F62672C8EAB3002"), "rb.vsEbAlpha.sup"
    gStr.Add U("80CC666F900F660E5EA6"), "rb.vsEbAlpha.tip"
    gStr.Add U("81EA8A020020004400500049"), "rb.vsEbDpi"
    gStr.Add U("0031FF5E003300300030003000204E4B95937684657465783002300C0050004E0047002089E367905EA6300D9078300C81EA8A02300D6642751F6548300289E367905EA68D8A9AD86A9468488D8A59273002"), "rb.vsEbDpi.sup"
    gStr.Add U("81EA8A0289E367905EA6"), "rb.vsEbDpi.tip"
    gStr.Add U("683C7DDA002878C50029"), "rb.vsEbGrid"
    gStr.Add U("683C7DDA80CC666F7684683C5B5059275C0FFF0C55AE4F4D70BA78C5FF1B6BCF002000350020683C4E00689D6DF182727DDA3002"), "rb.vsEbGrid.sup"
    gStr.Add U("683C7DDA95938DDD"), "rb.vsEbGrid.tip"
    gStr.Add U("908A754C002878C50029"), "rb.vsEbMargin"
    gStr.Add U("57165F6256DB54687559767D76845BEC5EA6FF0C55AE4F4D70BA78C5FF080031002078C50020224800200030002E0033003500206BEB7C73FF093002003000208868793A7DCA8CBC57165F623002"), "rb.vsEbMargin.sup"
    gStr.Add U("908A754C"), "rb.vsEbMargin.tip"
    gStr.Add U("81EA8A0280CC666F8272"), "rb.vsEbRgb.tip"
    gStr.Add U("81EA8A02984F8272"), "rb.vsGalColor"
    s = ""
    s = s & U("76EE524D81EA8A0280CC666F8272FF1A007B0030007D")
    s = s & vbCrLf
    s = s & U("5F9E4E3B984C984F827230016A196E968272300167008FD14F7F75287684984F82724E2D907864C7FF0C62167528300C51764ED6984F8272300D300C53D682725668300DFF1B4E5F53EF4EE557280020005200470042002068464E2D8F38516500200052002C0047002C0042FF0859820020003200350035002C003200340038002C003200330035FF0962165341516D90324F4DFF08598200200023004600460046003800450042FF093002907882725F8C80CC666F81EA52D5520763DB70BA300C81EA8A02984F8272300D3002")
    gStr.Add s, "rb.vsGalColor.sup"
    gStr.Add U("8D0A52A94F5C8005"), "rb.vsGalDonate"
    gStr.Add U("77E253706C384E45514D8CBB3001958B6E9030025982679C5B835E6B52304E864F60FF0C53EF4EE57528652F4ED85BF662165FAE4FE1638378BC8ACB4F5C8005559D676F54965561FF085B8C516881EA9858FF0C4E0D5F7197FF4EFB4F55529F80FDFF093002"), "rb.vsGalDonate.sup"
    gStr.Add U("95DC65BC"), "rb.vsGrpAbout"
    gStr.Add U("80CC666F"), "rb.vsGrpBackground"
    gStr.Add U("532F51FA907853D657165F62"), "rb.vsGrpExport"
    gStr.Add U("532F51FA8A2D5B9A"), "rb.vsGrpSettings"
    gStr.Add U("62796B2182075DE55177"), "rb.vsGrpTools"
    gStr.Add U("80CC666FFF1A007B0030007D"), "rb.vsMnuBg"
    s = ""
    s = s & U("76EE524D80CC666FFF1A007B0030007D")
    s = s & vbCrLf
    s = s & U("532F51FA6642588A572857165F624E0B9762768480CC666FFF1A900F660E3001767D827230017C7382723001683C7DDAFF085EA76A197D19FF09621681EA8A02984F82723002300C81EA8A02984F8272300D4E2D53EF5F9E4E3B984C984F827230016A196E968272300167008FD14F7F75287684984F82724E2D907864C7FF0C62167528300C51764ED6984F8272300D300C53D682725668300D30029EDE9078984F827265B9584A53EF76F463A5958B555F984F82729762677F3002")
    gStr.Add s, "rb.vsMnuBg.sup"
    gStr.Add U("80CC666F"), "rb.vsMnuBg.tip"
    gStr.Add U("0057006F007200640020002F00200057005000530020002F00200050006F0077006500720050006F0069006E00740020901A75287684541191CF5716FF0C8AD6658763D25716999690783002"), "rb.vsMnuEMF.desc"
    gStr.Add U("0067007A00690070002058D37E2E768400200045004D0046FF0C51675BB976F8540CFF0C9AD47A4D66F45C0F3002"), "rb.vsMnuEMZ.desc"
    gStr.Add U("76EE524DFF1A007B0030007D3002520763DB77E2537076844ECB97628A9E8A00FF087C219AD44E2D658730017E419AD44E2D658730010045006E0067006C006900730068300165E5672C8A9EFF09FF0C7ACB5373751F65483002"), "rb.vsMnuLang.sup"
End Sub

Private Sub LoadStrings_zh_TW_5()
    Dim s As String
    gStr.Add U("4ECB97628A9E8A000020002F0020004C0061006E00670075006100670065"), "rb.vsMnuLang.tip"
    gStr.Add U("9801976259275C0F820757165F624E0081F4FF0C906954080020004C00610054006500580020548C62957A3F7CFB7D713002"), "rb.vsMnuPDF.desc"
    gStr.Add U("0031003500300020002F00200033003000300020002F002000360030003000200044005000490020621681EA8A02FF0867009AD8002000330030003000300020004400500049FF093002"), "rb.vsMnuPNG.desc"
    gStr.Add U("7DB2980130010049006C006C007500730074007200610074006F007230010049006E006B007300630061007000650020901A7528FF1B0050006F0077006500720050006F0069006E007400200032003000310033002053CA4EE54E0A90FD80FD532F51FA3002"), "rb.vsMnuSVG.desc"
    gStr.Add U("65874EF6"), "rb.vsMnuSepDoc"
    gStr.Add U("9EDE96635716"), "rb.vsMnuSepRaster"
    gStr.Add U("541191CF5716"), "rb.vsMnuSepVector"
    gStr.Add U("0056006900730069006F0020539F751F5F6272C0FF0C65875B5753EF76F463A57DE88F2FFF1B712197005B8988DD00200056006900730069006F3002"), "rb.vsMnuVSDX.desc"
    gStr.Add U("76F85BB980018EDF9AD4FF1B4E0D652F63F4900F660E548C6F388B8A3002"), "rb.vsMnuWMF.desc"
    gStr.Add U("77E25370"), "rb.vsTab"
    gStr.Add U("76EE524D8A2D5B9A"), "set.current"
    gStr.Add U("FF08672A8A2D5B9AFF1A4F7F75287C215831624057288CC76599593EFF1B672A51325B586216572896F27AEF66424F7F7528684C9762FF09"), "set.folderNone"
    gStr.Add U("907864C7300C81EA52D551325B58300D4F7F7528768498108A2D8CC76599593E"), "set.pickDefaultFolder"
    gStr.Add U("5206679057165F627D5069CB"), "step.analyze"
    gStr.Add U("572881E8664262955F7172474E0A5408621057165F62300180CC666F8207908A754C"), "step.compose"
    gStr.Add U("540862105B8C6210"), "step.composed"
    gStr.Add U("5B8C6210FF1A007B0030007D"), "step.doneFile"
    gStr.Add U("4F7F752800200049006E006B007300630061007000650020002F0020004C0069006200720065004F006600660069006300650020628A00200045004D004600208F4963DB70BA0020005300560047"), "step.emf2svg"
    gStr.Add U("58D37E2E70BA00200045004D005AFF080067007A00690070FF09"), "step.emz"
    gStr.Add U("572862955F7172474E0A9EDE9078898162FE53D67684984F8272FF0800360030002079D2516767096548FF09"), "step.eyedropper"
    gStr.Add U("6AA267E54E264FEE5FA90020005300560047"), "step.fixSvg"
    gStr.Add U("0049006E006B0073006300610070006500206B635728628A00200045004D004600208F4963DB70BA0020005300560047"), "step.inkscape"
    gStr.Add U("004C0069006200720065004F0066006600690063006500206B635728628A00200045004D004600208F4963DB70BA0020005300560047"), "step.libreoffice"
    gStr.Add U("5EFA7ACB820757165F62540C6A2359275C0F768481E8664265874EF6"), "step.pdfDoc"
    gStr.Add U("4EE5541191CF57165F625F0F8CBC4E0A"), "step.pdfPaste"
    gStr.Add U("0050006F0077006500720050006F0069006E007400206B6357287522751F0020005000440046"), "step.pdfWrite"
    gStr.Add U("6E3267D300200050004E0047FF1A007B0030007D002000D70020007B0031007D0020756B7D20FF08007B0032007D0020004400500049FF09"), "step.png"
    gStr.Add U("0050006F0077006500720050006F0069006E007400206B6357287522751F00200045004D00460020541191CF5716"), "step.ppEmf"
    gStr.Add U("0050006F0077006500720050006F0069006E007400206B635728532F51FA0020005300560047"), "step.ppSvg"
    gStr.Add U("0050006F0077006500720050006F0069006E007400206B6357287522751F00200057004D00460020571651436A946848"), "step.ppWmf"
    gStr.Add U("6E96509957165F62FF08007B0030007D0020500B72694EF6FF09"), "step.prepare"
    gStr.Add U("7522751F00200045004D00460020541191CF5716FF08571672476A215F0FFF09"), "step.vsdxPicture"
    gStr.Add U("7522751F00200056006900730069006F00205F6272C00020007B0030007D0020002F0020007B0031007DFF1A007B0032007D"), "step.vsdxShape"
    gStr.Add U("5BEB51FA00200056006900730069006F00206A946848FF08007B0030007D0020500B539F751F5F6272C0FF0C007B0031007D00205F3557167247FF09"), "step.vsdxWrite"
    gStr.Add U("5BEB51FA00200056006900730069006F00206A946848"), "step.vsdxWrite0"
    gStr.Add U("900F904E002000580050005300207522751F0020005300560047FF1A007B0030007D"), "step.xps"
    gStr.Add U("6E96509981E8664262955F717247FF0896B185CF6BCD7248548C80CC666FFF09"), "step.xpsPrep"
    gStr.Add U("0050006F0077006500720050006F0069006E007400206B6357285217537070BA0020005800500053"), "step.xpsPrint"
    gStr.Add U("6253530500200056005300440058"), "step.zip"
    gStr.Add U("FF0C900F660E5EA60020007B0030007D0025"), "sum.alpha"
    gStr.Add U("80CC666FFF1A007B0030007D"), "sum.bg"
    gStr.Add U("0050004E0047002089E367905EA6FF1A007B0030007D0020004400500049"), "sum.dpi"
    gStr.Add U("98108A2D683C5F0FFF1A007B0030007D"), "sum.format"
    gStr.Add U("630976EE524D8A2D5B9A532F51FA907853D657165F62"), "sum.head"
    gStr.Add U("4ECB97628A9E8A00FF1A007B0030007D"), "sum.lang"
    gStr.Add U("908A754CFF1A007B0030007D002078C5"), "sum.margin"
    gStr.Add U("591A500B57165F62FF1A007B0030007D"), "sum.mode"
    gStr.Add U("51325B5865B95F0FFF1A007B0030007D"), "sum.naming"
    gStr.Add U("0056005300440058FF1A0056006900730069006F0020539F751F5F6272C0FF0865875B5753EF7DE88F2FFF09"), "sum.vsdxNative"
    gStr.Add U("0056005300440058FF1A5D4C516500200045004D00460020541191CF57167247"), "sum.vsdxPicture"
End Sub

Private Sub LoadStrings_zh_TW_6()
    Dim s As String
    gStr.Add U("8F4963DB5B8C6210"), "xps.done"
    gStr.Add U("8F4963DB57165F6251437D200020007B0030007D0020002F0020007B0031007D"), "xps.elem"
    gStr.Add U("5D4C51655B57578BFF087B2C0020007B0030007D0020500BFF09"), "xps.font"
    gStr.Add U("8B8053D6002000580050005300206A946848"), "xps.open"
    gStr.Add U("5206679098019762FF1A51710020007B0030007D0020500B57165F6251437D20"), "xps.page"
    gStr.Add U("555F52D58F4963DB7A0B5F0F"), "xps.start"
    gStr.Add U("5BEB51FA002000530056004700206A946848"), "xps.write"
End Sub

Private Sub LoadStrings_en_1()
    Dim s As String
    gStr.Add "Language: {0} (change it with the language menu in the About group)", "about.lang"
    s = ""
    s = s & "VecStamp  v{0}"
    s = s & vbCrLf
    s = s & U("00530065006C00650063007400200061006E00640020007300740061006D0070002020140020006500780070006F00720074002000730068006100700065007300200074006F00200045004D00460020002F00200053005600470020002F00200050004400460020002F002000560053004400580020002F00200050004E004700200069006E0020006F006E006500200063006C00690063006B")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & "Author: {1}"
    s = s & vbCrLf
    s = s & "E-mail: {2}"
    s = s & vbCrLf
    s = s & "Home page: {3}"
    s = s & vbCrLf
    s = s & "Licence: MIT License"
    s = s & vbCrLf
    s = s & "Copyright (c) {4} {1}"
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & "VecStamp is and will stay free and open source: no paid features, no ads, no data collection, all source code on GitHub."
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & "If it saves you time, you can buy the author a coffee with ""Support the author"" in the About group (WeChat Pay or Alipay). It is entirely voluntary and nothing depends on it; support goes into maintai"
    s = s & "ning and improving VecStamp."
    gStr.Add s, "about.text"
    gStr.Add "About {0}", "about.title"
    gStr.Add "VecStamp", "app.title"
    gStr.Add "Beige", "bg.beige"
    gStr.Add "Custom colour", "bg.custom"
    gStr.Add "Custom colour {0}", "bg.customValue"
    gStr.Add "Grid", "bg.grid"
    gStr.Add "Grid ({0} pt spacing)", "bg.gridValue"
    gStr.Add "Transparent", "bg.none"
    gStr.Add "White", "bg.white"
    gStr.Add "Cancel", "btn.cancel"
    gStr.Add "OK", "btn.ok"
    gStr.Add "Colour examples: 255,248,235 or #FFF8EB", "col.bad"
    gStr.Add "Darker {0}%", "col.darker"
    gStr.Add "Click a swatch to choose it, double-click to confirm; or type R,G,B or #RRGGBB.", "col.hint"
    gStr.Add "Lighter {0}%", "col.lighter"
    gStr.Add "Type a number to choose a colour, or R,G,B / #RRGGBB:", "col.listPrompt"
    gStr.Add "Recent colours", "col.recent"
    gStr.Add "Recent colour: {0}", "col.recentItem"
    gStr.Add "Standard colours", "col.standard"
    gStr.Add "Standard colour: {0}", "col.stdItem"
    gStr.Add "Dark Red|Red|Orange|Yellow|Light Green|Green|Light Blue|Blue|Dark Blue|Purple", "col.stdNames"
    gStr.Add U("00530079007300740065006D00200063006F006C006F0075007200732026"), "col.system"
    gStr.Add "System colour", "col.systemPicked"
    gStr.Add "Theme colours", "col.theme"
    gStr.Add "White, Background 1|Black, Text 1|Background 2|Text 2|Accent 1|Accent 2|Accent 3|Accent 4|Accent 5|Accent 6", "col.themeNames"
    gStr.Add "{0}, {1}", "col.tint"
    gStr.Add "Custom colour", "col.title"
    gStr.Add "Colour", "col.value"
    gStr.Add "{0} ({1})", "col.withValue"
    gStr.Add "File name (without extension):", "dlg.fileName"
    gStr.Add "Choose the export folder", "dlg.pickFolder"
    gStr.Add "Choose the folder for the slide files (format: {0})", "dlg.pickFolderBatch"
    gStr.Add "Choose the export folder (one file per shape)", "dlg.pickFolderEach"
    gStr.Add "Export as {0}", "dlg.saveAs"
    gStr.Add "Export selected shapes", "dlg.saveTitle"
    gStr.Add "Alipay", "don.alipay"
    s = ""
    s = s & "Thank you for wanting to support VecStamp!"
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & "Scan the matching QR code in ""Support the author"" with {0}; any amount is welcome."
    s = s & vbCrLf
    s = s & "Support is entirely voluntary: every feature of VecStamp is and stays free and open source."
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & "A star for the project on GitHub, or recommending it to colleagues, helps a lot too."
    gStr.Add s, "don.thanks"
    gStr.Add "Support", "don.title"
    gStr.Add "WeChat", "don.wechat"
    gStr.Add "Custom", "dpi.custom"
    gStr.Add "The selected shapes could not be copied through the clipboard (another program may be using it). Please try again.", "err.clipboard"
    gStr.Add "EMZ compression failed.", "err.emz"
    gStr.Add "The selected objects cannot be exported together.", "err.group"
    gStr.Add "The macOS helper script {0} was not found. Please run the macOS installer again.", "err.macHelper"
    gStr.Add "(PowerPoint for Mac 16.82 or later is needed to export shapes)", "err.macVersion"
    gStr.Add "The file could not be written: {0}", "err.macWrite"
    gStr.Add "No file was written: {0}", "err.noFile"
    gStr.Add "PDF export failed: {0}", "err.pdf"
End Sub

Private Sub LoadStrings_en_2()
    Dim s As String
    gStr.Add "Packing the VSDX failed.", "err.zip"
    gStr.Add "Packing the VSDX failed (it needs the Windows PowerShell that comes with Windows).", "err.zipPs"
    s = ""
    s = s & "How would you like to send feedback?"
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & "[Yes] E-mail the author ({0})"
    s = s & vbCrLf
    s = s & "[No] Open a GitHub issue (needs a GitHub account; issues are public)"
    s = s & vbCrLf
    s = s & "[Cancel] Back"
    gStr.Add s, "fb.ask"
    s = ""
    s = s & "**Problem / suggestion:**"
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & "**Steps to reproduce:**"
    s = s & vbCrLf
    s = s & "1. "
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & "**Environment:**"
    s = s & vbCrLf
    s = s & "```"
    s = s & vbCrLf
    s = s & "{0}"
    s = s & vbCrLf
    s = s & "```"
    s = s & vbCrLf
    gStr.Add s, "fb.issueBody"
    gStr.Add "[Feedback] ", "fb.issueTitle"
    s = ""
    s = s & "The author's address has been copied to the clipboard: {0}"
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & "You can paste it as the recipient in any mail service (web mail, Outlook, ...)."
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & "Open your default mail program now?"
    s = s & vbCrLf
    s = s & "(If none is set up, choose No and use your web mail.)"
    gStr.Add s, "fb.mailCopied"
    s = ""
    s = s & "The author's address: {0}"
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & "You can write to it from any mail service (web mail, Outlook, ...)."
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & "Open your default mail program now?"
    s = s & vbCrLf
    s = s & "(If none is set up, choose No and use your web mail.)"
    gStr.Add s, "fb.mailPlain"
    gStr.Add "Feedback by e-mail", "fb.mailTitle"
    s = ""
    s = s & "No browser could be opened. Please visit:"
    s = s & vbCrLf
    s = s & "{0}"
    gStr.Add s, "fb.noBrowser"
    gStr.Add "No mail program could be opened. Please write to: {0}", "fb.noMailApp"
    gStr.Add "VecStamp {0} feedback", "fb.subject"
    gStr.Add "Feedback", "fb.title"
    gStr.Add "slide{0}", "file.page"
    gStr.Add "Finder", "fm.mac"
    gStr.Add "File Explorer", "fm.win"
    gStr.Add "EMF Enhanced Metafile", "fmt.emf"
    gStr.Add "EMZ Compressed Enhanced Metafile", "fmt.emz"
    gStr.Add "PDF Vector document", "fmt.pdf"
    gStr.Add "PNG High-resolution bitmap", "fmt.png"
    gStr.Add "SVG Scalable Vector Graphics", "fmt.svg"
    gStr.Add "VSDX Visio drawing", "fmt.vsdx"
    gStr.Add "WMF Windows Metafile", "fmt.wmf"
    s = ""
    s = s & U("005600650063005300740061006D00700020201300200071007500690063006B002000670075006900640065")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("2460002000530065006C0065006300740020006F006E00650020006F00720020006D006F00720065002000730068006100700065007300200061006E006400200063006C00690063006B0020004500780070006F00720074002C0020006F00720020006F006E00650020006F0066002000740068006500200066006F0072006D0061007400200062007500740074006F006E0073002E002000570069007400680020006E006F007400680069006E0067002000730065006C0065006300740065006400200079006F0075002000630061006E0020006500780070006F007200740020007400680065002000770068006F006C006500200073006C006900640065002E")
    s = s & vbCrLf
    s = s & U("24610020004500780070006F0072007400200065007600650072007900200073006C006900640065003A0020006F006E0065002000660069006C0065002000700065007200200073006C006900640065003B002000730065006C00650063007400200073006C006900640065007300200069006E00200074006800650020007400680075006D0062006E00610069006C002000700061006E006500200074006F0020006500780070006F007200740020006F006E006C0079002000740068006F00730065002E")
    s = s & vbCrLf
    s = s & U("246200200045004D00460020007000690063007400750072006500200063006F00700079003A00200070006C00610063006500730020006100200076006500630074006F0072002000700069006300740075007200650020006F00660020007400680065002000730065006C0065006300740069006F006E0020006E00650078007400200074006F002000690074003B002000740068006500200073006800610070006500730020007300740061007900200075006E006300680061006E006700650064002E")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("00B700200045004D00460020002F00200045004D005A0020002F00200057004D00460020002F0020005300560047003A00200076006500630074006F0072002000700069006300740075007200650073002E00200045007600650072007900200050006F0077006500720050006F0069006E0074002000320030003100330020006F00720020006C0061007400650072002000630061006E0020006500780070006F0072007400200053005600470020002800760065007200730069006F006E00730020006200650066006F00720065002000320033003000320020007400680072006F00750067006800200074006800650020006200750069006C0074002D0069006E002000580050005300200063006F006E00760065007200730069006F006E0029002E")
    s = s & vbCrLf
    s = s & U("00B70020005000440046003A0020006100200076006500630074006F00720020005000440046002000650078006100630074006C00790020006100730020006C00610072006700650020006100730020007400680065002000640072006100770069006E0067002C00200069006400650061006C00200066006F00720020004C0061005400650058002E")
    s = s & vbCrLf
    s = s & U("00B700200056005300440058003A0020006100200056006900730069006F002000660069006C0065002000770069007400680020006E0061007400690076006500200056006900730069006F002000730068006100700065007300200061006E00640020006500640069007400610062006C006500200074006500780074003B00200056006900730069006F0020006900730020006E006F00740020006E00650065006400650064002E")
    s = s & vbCrLf
    s = s & U("00B700200050004E0047003A00200031003500300020002F00200033003000300020002F002000360030003000200044005000490020006F007200200063007500730074006F006D002000280075007000200074006F0020007B0030007D00200044005000490029002E")
    s = s & vbCrLf
    s = s & U("00B70020004200610063006B00670072006F0075006E0064003A00200069006E00200074006800650020004200610063006B00670072006F0075006E00640020006D0065006E0075002000630068006F006F007300650020007400720061006E00730070006100720065006E0074002C002000770068006900740065002C002000620065006900670065002C002000670072006900640020006F00720020006100200063007500730074006F006D00200063006F006C006F0075007200200028007400680065006D00650020002F0020007300740061006E006400610072006400200063006F006C006F007500720073002C0020004D006F0072006500200043006F006C006F007500720073002C002000450079006500640072006F0070007000650072002C00200052004700420029002C002000770069007400680020007400720061006E00730070006100720065006E0063007900200061006E00640020006D0061007200670069006E002E")
    s = s & vbCrLf
    s = s & U("00B700200041002000700072006F00670072006500730073002000770069006E0064006F0077002000730068006F007700730020007400680065002000630075007200720065006E00740020007300740065007000200061006E0064002000700065007200630065006E0074006100670065003B002000700072006500730073002000450073006300200074006F002000630061006E00630065006C002E")
    s = s & vbCrLf
    s = s & U("00B70020004C0061006E00670075006100670065003A00200073007700690074006300680020006200650074007700650065006E00207B804F534E2D6587002C00207E419AD44E2D6587002C00200045006E0067006C00690073006800200061006E0064002065E5672C8A9E00200069006E0020007400680065002000410062006F00750074002000670072006F00750070003B002000690074002000740061006B0065007300200065006600660065006300740020006100740020006F006E00630065002E")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & "Combined exports, backgrounds, PDF and slide-by-slide export briefly use the clipboard and a temporary slide at the end (removed automatically)."
    gStr.Add s, "help.ppt"
    gStr.Add "One file for all", "mode.merge"
    gStr.Add "One file per shape", "mode.split"
    gStr.Add "Reset all export settings to their defaults?", "msg.askReset"
    s = ""
    s = s & "Which slides should be exported?"
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & "Yes: only the {0} selected slides"
    s = s & vbCrLf
    s = s & "No: all {1} slides"
    gStr.Add s, "msg.askWhich"
    s = ""
    s = s & "No shapes are selected."
    s = s & vbCrLf
    s = s & "Export all shapes on this slide?"
    gStr.Add s, "msg.askWholeSlide"
    gStr.Add "Please enter a background transparency from 0 to 100 (%).", "msg.badAlpha"
    gStr.Add "Please enter a resolution from 1 to {0} (DPI).", "msg.badDpi"
    gStr.Add "Please enter a grid spacing from 2 to 200 (points).", "msg.badGrid"
    gStr.Add "Please enter a margin from 0 to 500 (points).", "msg.badMargin"
    gStr.Add "Exporting the slides failed: {0}", "msg.batchFailed"
    gStr.Add U("004100740020007B0030007D0020004400500049002000740068006900730020007000690063007400750072006500200077006F0075006C0064002000620065002000610062006F007500740020007B0031007D002000D70020007B0032007D00200070006900780065006C007300200028007B0033007D0020006D0065006700610070006900780065006C0073002900200061006E00640020006D0061007900200062006500200073006C006F00770020006F0072002000720075006E0020006F007500740020006F00660020006D0065006D006F00720079002E00200043006F006E00740069006E0075006500200061006E0079007700610079003F"), "msg.bigPng"
    gStr.Add "Export cancelled.", "msg.cancelled"
    gStr.Add "Export cancelled. The {0} file(s) already finished were kept:", "msg.cancelledPartial"
    s = ""
    s = s & "The colour picker could not be opened: {0}"
    s = s & vbCrLf
    s = s & "You can also type the colour into the RGB box."
    gStr.Add s, "msg.colorFailed"
    gStr.Add "Export failed: {0}", "msg.exportFailed"
    gStr.Add "{0} file(s) exported:", "msg.exported"
    s = ""
    s = s & "The eyedropper needs a slide open in Normal view."
    s = s & vbCrLf
    s = s & "Switch to Normal view, click Eyedropper again, then click the colour on the slide."
    gStr.Add s, "msg.eyedropperView"
    gStr.Add "Project page: {0}", "msg.home"
    gStr.Add "Nothing has been exported yet. After the first export this opens the most recent folder.", "msg.noFolderYet"
    gStr.Add "There are no shapes to export on this slide.", "msg.noShapesSlide"
    gStr.Add "There are no shapes to export on the chosen slides.", "msg.noShapesSlides"
    gStr.Add "The presentation has no slides.", "msg.noSlides"
    s = ""
    s = s & "No SVG could be created: this PowerPoint cannot export SVG itself, the built-in conversion did not succeed (it needs the Windows PowerShell that comes with Windows), and neither Inkscape nor LibreOffi"
    s = s & "ce is installed."
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & "Export EMF instead? (You can also export a PDF and convert it with another program.)"
    gStr.Add s, "msg.noSvg"
    gStr.Add "Please select the shapes to export in Normal view.", "msg.normalView"
    gStr.Add "Please open a presentation and select the shapes to export.", "msg.openAndSelect"
    gStr.Add "Please open a presentation first.", "msg.openPres"
    gStr.Add "Could not create the EMF picture copy: {0}", "msg.pictureFailed"
End Sub

Private Sub LoadStrings_en_3()
    Dim s As String
    gStr.Add "Please select shapes first.", "msg.selectFirst"
    gStr.Add "{0} empty slide(s) skipped.", "msg.skipped"
    s = ""
    s = s & "The built-in SVG conversion did not finish within {0} seconds and was stopped."
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & "This can happen with very complex drawings (lots of text, pictures or points)."
    s = s & vbCrLf
    s = s & "Convert with Inkscape / LibreOffice instead (if installed), or export EMF?"
    gStr.Add s, "msg.svgTimeout"
    gStr.Add "Ask where to save", "naming.ask"
    gStr.Add "Save to the default folder", "naming.auto"
    gStr.Add "Exporting slides ({0} slides, {1})", "prog.batch"
    gStr.Add " ({0} s elapsed)", "prog.elapsed"
    gStr.Add "Press Esc to cancel", "prog.escHint"
    gStr.Add "Exporting {0}", "prog.export1"
    gStr.Add "Exporting {0} {1} files", "prog.exportN"
    gStr.Add "Eyedropper", "prog.eyedropper"
    gStr.Add "Shape {0} of {1}: {2}", "prog.item"
    gStr.Add "Slide {0} of {1} (slide {2})", "prog.page"
    gStr.Add U("0050007200650070006100720069006E00672026"), "prog.preparing"
    gStr.Add "Beige", "rb.vsBgBeige"
    gStr.Add U("0043007500730074006F006D00200063006F006C006F007500722026"), "rb.vsBgCustomMac"
    gStr.Add "Grid (graph paper)", "rb.vsBgGrid"
    gStr.Add "Transparent", "rb.vsBgNone"
    gStr.Add "Custom colour", "rb.vsBgSepCustom"
    gStr.Add "White", "rb.vsBgWhite"
    gStr.Add "About VecStamp", "rb.vsBtnAbout"
    gStr.Add U("00560065007200730069006F006E002C00200061007500740068006F0072002C0020004D004900540020006C006900630065006E00630065002020130020005600650063005300740061006D007000200069007300200061006E00640020007300740061007900730020006600720065006500200061006E00640020006F00700065006E00200073006F0075007200630065002E"), "rb.vsBtnAbout.sup"
    gStr.Add "About VecStamp", "rb.vsBtnAbout.tip"
    gStr.Add "Export every slide", "rb.vsBtnBatch"
    gStr.Add "Exports every slide as one file (in the default format). With several slides selected in the thumbnail pane only those are exported; empty slides and empty title placeholders are skipped.", "rb.vsBtnBatch.sup"
    gStr.Add U("00410062006F0075007400200073007500700070006F007200742026"), "rb.vsBtnDonateInfo"
    gStr.Add U("0045006E00680061006E0063006500640020004D00650074006100660069006C006500200066006F007200200057006F007200640020002F00200057005000530020002F00200050006F0077006500720050006F0069006E0074002020130020007400680065002000660069007200730074002000630068006F00690063006500200066006F00720020007000610070006500720073002E"), "rb.vsBtnEMF.sup"
    gStr.Add "Export as EMF", "rb.vsBtnEMF.tip"
    gStr.Add "EMF compressed with gzip: the same content in a smaller file.", "rb.vsBtnEMZ.sup"
    gStr.Add "Export as EMZ", "rb.vsBtnEMZ.tip"
    gStr.Add "Export {0}", "rb.vsBtnExport"
    gStr.Add "Export the selected shapes", "rb.vsBtnExport.tip"
    gStr.Add "Eyedropper", "rb.vsBtnEyedropper"
    gStr.Add "Feedback", "rb.vsBtnFeedback"
    gStr.Add "E-mail the author or open an issue on GitHub.", "rb.vsBtnFeedback.sup"
    gStr.Add "Report a problem or suggest an idea", "rb.vsBtnFeedback.tip"
    gStr.Add U("00440065006600610075006C007400200066006F006C0064006500722026"), "rb.vsBtnFolder"
    s = ""
    s = s & "The folder used for automatic saving and Export every slide:"
    s = s & vbCrLf
    s = s & "{0}"
    gStr.Add s, "rb.vsBtnFolder.sup"
    gStr.Add "Help", "rb.vsBtnHelp"
    gStr.Add "A short guide to the buttons and settings.", "rb.vsBtnHelp.sup"
    gStr.Add "Project page", "rb.vsBtnHome"
    gStr.Add "Opens the VecStamp page on GitHub: new versions, source code and change log.", "rb.vsBtnHome.sup"
    gStr.Add "Shows all current export settings.", "rb.vsBtnLauncher.sup"
    gStr.Add "Show the current settings", "rb.vsBtnLauncher.tip"
    gStr.Add U("004D006F0072006500200043006F006C006F0075007200732026"), "rb.vsBtnMoreColors"
    gStr.Add "Open export folder", "rb.vsBtnOpenFolder"
    gStr.Add "Opens the folder of the last export in {0}.", "rb.vsBtnOpenFolder.sup"
    gStr.Add "A vector PDF exactly as large as the drawing.", "rb.vsBtnPDF.sup"
    gStr.Add "Export as PDF", "rb.vsBtnPDF.tip"
    gStr.Add U("0048006900670068002D007200650073006F006C007500740069006F006E0020006200690074006D00610070003B002000630068006F006F0073006500200074006800650020007200650073006F006C007500740069006F006E00200075006E0064006500720020004500780070006F00720074002000730065007400740069006E006700730020219200200050004E00470020007200650073006F006C007500740069006F006E002E"), "rb.vsBtnPNG.sup"
End Sub

Private Sub LoadStrings_en_4()
    Dim s As String
    gStr.Add "Export as PNG", "rb.vsBtnPNG.tip"
    gStr.Add "EMF picture copy", "rb.vsBtnPicture"
    gStr.Add "Places a vector picture of the selection next to it to ""freeze"" the layout; the shapes stay unchanged.", "rb.vsBtnPicture.sup"
    gStr.Add "Reset settings", "rb.vsBtnReset"
    gStr.Add "Resets all export and background settings to their defaults (the language is kept).", "rb.vsBtnReset.sup"
    s = ""
    s = s & "Scalable Vector Graphics that any browser can open. Works with every PowerPoint 2013 or later: newer versions export it directly, older ones through the built-in XPS conversion (text stays text); the "
    s = s & "file is checked afterwards."
    gStr.Add s, "rb.vsBtnSVG.sup"
    gStr.Add "Export as SVG", "rb.vsBtnSVG.tip"
    s = ""
    s = s & U("0056006900730069006F002000640072006100770069006E0067003A0020006500760065007200790020007300680061007000650020006200650063006F006D00650073002000610020006E0061007400690076006500200056006900730069006F0020007300680061007000650020002800670065006F006D0065007400720079002C00200063006F006C006F007500720073002C0020006C0069006E00650073002C0020006100720072006F00770073002C002000670072006F0075007000730029002000770069007400680020006500640069007400610062006C0065002000740065007800740020201300200056006900730069006F0020006900730020006E006F00740020006E00650065006400650064002E002000500069006300740075007200650073002C002000630068006100720074007300200061006E00640020006F00740068006500720020006F0062006A006500630074007300200077006900740068006F007500740020006100200056006900730069006F0020006500710075006900760061006C0065")
    s = s & "nt are embedded as vector pictures."
    gStr.Add s, "rb.vsBtnVSDX.sup"
    gStr.Add "Export as VSDX", "rb.vsBtnVSDX.tip"
    gStr.Add "Open folder when done", "rb.vsChkOpen"
    gStr.Add "Shows the exported file in File Explorer / Finder when done.", "rb.vsChkOpen.sup"
    gStr.Add "VSDX native shapes", "rb.vsChkVsdx"
    s = ""
    s = s & "On: the shapes in the VSDX are native Visio shapes with editable text (Visio is not needed)."
    s = s & vbCrLf
    s = s & "Off: the whole drawing is embedded as one EMF vector picture and looks exactly as in PowerPoint."
    gStr.Add s, "rb.vsChkVsdx.sup"
    gStr.Add "PNG resolution", "rb.vsDdDpi"
    gStr.Add "Only for PNG. Journals usually ask for 300 DPI or more, line art 600 DPI; custom values up to 3000 DPI are possible.", "rb.vsDdDpi.sup"
    gStr.Add "Default format", "rb.vsDdFormat"
    gStr.Add "The format used by Export and Export every slide (it can also be changed in the save dialog).", "rb.vsDdFormat.sup"
    gStr.Add "Several shapes", "rb.vsDdMode"
    gStr.Add "One file for all, or one file per shape (named after the shapes in the Selection Pane).", "rb.vsDdMode.sup"
    gStr.Add "With several shapes selected", "rb.vsDdMode.tip"
    gStr.Add "Saving", "rb.vsDdNaming"
    gStr.Add "Ask every time, or save straight to the default folder (file names get the time).", "rb.vsDdNaming.sup"
    gStr.Add "Transp. %", "rb.vsEbAlpha"
    gStr.Add "0 = opaque, 100 = fully transparent; only the background, not the shapes.", "rb.vsEbAlpha.sup"
    gStr.Add "Background transparency", "rb.vsEbAlpha.tip"
    gStr.Add "Custom DPI", "rb.vsEbDpi"
    gStr.Add "A whole number from 1 to 3000, used when PNG resolution is Custom. Higher means larger files.", "rb.vsEbDpi.sup"
    gStr.Add "Custom resolution", "rb.vsEbDpi.tip"
    gStr.Add "Grid (pt)", "rb.vsEbGrid"
    gStr.Add "Size of the grid squares in points; every 5th line is darker.", "rb.vsEbGrid.sup"
    gStr.Add "Grid spacing", "rb.vsEbGrid.tip"
    gStr.Add "Margin (pt)", "rb.vsEbMargin"
    gStr.Add U("00530070006100630065002000610072006F0075006E00640020007400680065002000730068006100700065007300200069006E00200070006F0069006E007400730020002800310020007000740020224800200030002E003300350020006D006D0029002E002000300020003D002000740069006700680074002E"), "rb.vsEbMargin.sup"
    gStr.Add "Margin", "rb.vsEbMargin.tip"
    gStr.Add "Custom background colour", "rb.vsEbRgb.tip"
    gStr.Add "Custom colour", "rb.vsGalColor"
    s = ""
    s = s & "Current custom background colour: {0}"
    s = s & vbCrLf
    s = s & "Pick a theme, standard or recent colour, or use More Colours or the Eyedropper; you can also type R,G,B (e.g. 255,248,235) or hex (e.g. #FFF8EB) into the RGB box. Picking a colour switches the backgro"
    s = s & "und to Custom colour."
    gStr.Add s, "rb.vsGalColor.sup"
    gStr.Add "Support the author", "rb.vsGalDonate"
    gStr.Add "VecStamp is free and open source. If it helps you, you can buy the author a coffee with Alipay or WeChat Pay (entirely voluntary; nothing depends on it).", "rb.vsGalDonate.sup"
    gStr.Add "About", "rb.vsGrpAbout"
    gStr.Add "Background", "rb.vsGrpBackground"
    gStr.Add "Export selection", "rb.vsGrpExport"
    gStr.Add "Export settings", "rb.vsGrpSettings"
    gStr.Add "Batch & tools", "rb.vsGrpTools"
    gStr.Add "Background: {0}", "rb.vsMnuBg"
    s = ""
    s = s & "Current background: {0}"
    s = s & vbCrLf
    s = s & "The background placed under the shapes when exporting: transparent, white, beige, grid (graph paper) or a custom colour. Under Custom colour pick a theme, standard or recent colour, or use More Colour"
    s = s & "s or the Eyedropper. Click the colour square to open the colour panel."
    gStr.Add s, "rb.vsMnuBg.sup"
    gStr.Add "Background", "rb.vsMnuBg.tip"
    gStr.Add U("0056006500630074006F00720020007000690063007400750072006500200066006F007200200057006F007200640020002F00200057005000530020002F00200050006F0077006500720050006F0069006E0074002020130020007400680065002000660069007200730074002000630068006F00690063006500200066006F00720020007000610070006500720073002E"), "rb.vsMnuEMF.desc"
    gStr.Add "EMF compressed with gzip: same content, smaller file.", "rb.vsMnuEMZ.desc"
    gStr.Add U("00430075007200720065006E0074003A0020007B0030007D002E00200053007700690074006300680065007300200074006800650020006C0061006E006700750061006700650020006F00660020005600650063005300740061006D0070002000287B804F534E2D6587002C00207E419AD44E2D6587002C00200045006E0067006C006900730068002C002065E5672C8A9E00290020006100740020006F006E00630065002E"), "rb.vsMnuLang.sup"
End Sub

Private Sub LoadStrings_en_5()
    Dim s As String
    gStr.Add "Language", "rb.vsMnuLang.tip"
    gStr.Add U("0050006100670065002000650078006100630074006C00790020006100730020006C00610072006700650020006100730020007400680065002000640072006100770069006E00670020201300200066006F00720020004C006100540065005800200061006E00640020007300750062006D0069007300730069006F006E002000730079007300740065006D0073002E"), "rb.vsMnuPDF.desc"
    gStr.Add "150 / 300 / 600 DPI or custom (up to 3000 DPI).", "rb.vsMnuPNG.desc"
    gStr.Add "For the web, Illustrator and Inkscape; works with every PowerPoint 2013 or later.", "rb.vsMnuSVG.desc"
    gStr.Add "Documents", "rb.vsMnuSepDoc"
    gStr.Add "Bitmaps", "rb.vsMnuSepRaster"
    gStr.Add "Vector pictures", "rb.vsMnuSepVector"
    gStr.Add "Native Visio shapes with editable text; Visio is not needed.", "rb.vsMnuVSDX.desc"
    gStr.Add "For older programs; no transparency or gradients.", "rb.vsMnuWMF.desc"
    gStr.Add "VecStamp", "rb.vsTab"
    gStr.Add "Current settings", "set.current"
    gStr.Add "(not set: the folder of the presentation; the desktop for unsaved or cloud presentations)", "set.folderNone"
    gStr.Add "Choose the default folder for automatic saving", "set.pickDefaultFolder"
    gStr.Add "Analysing the drawing", "step.analyze"
    gStr.Add "Combining shapes, background and margin on a temporary slide", "step.compose"
    gStr.Add "Combined", "step.composed"
    gStr.Add "Done: {0}", "step.doneFile"
    gStr.Add "Converting EMF to SVG with Inkscape / LibreOffice", "step.emf2svg"
    gStr.Add "Compressing to EMZ (gzip)", "step.emz"
    gStr.Add "Click the colour to pick on the slide (within 60 seconds)", "step.eyedropper"
    gStr.Add "Checking and repairing the SVG", "step.fixSvg"
    gStr.Add "Inkscape is converting the EMF to SVG", "step.inkscape"
    gStr.Add "LibreOffice is converting the EMF to SVG", "step.libreoffice"
    gStr.Add "Creating a temporary document of the drawing's size", "step.pdfDoc"
    gStr.Add "Pasting as a vector picture", "step.pdfPaste"
    gStr.Add "PowerPoint is creating the PDF", "step.pdfWrite"
    gStr.Add U("00520065006E0064006500720069006E0067002000740068006500200050004E0047003A0020007B0030007D002000D70020007B0031007D00200070006900780065006C007300200028007B0032007D00200044005000490029"), "step.png"
    gStr.Add "PowerPoint is creating the EMF vector picture", "step.ppEmf"
    gStr.Add "PowerPoint is exporting the SVG", "step.ppSvg"
    gStr.Add "PowerPoint is creating the WMF metafile", "step.ppWmf"
    gStr.Add "Preparing the shapes ({0} objects)", "step.prepare"
    gStr.Add "Creating the EMF vector picture (picture mode)", "step.vsdxPicture"
    gStr.Add "Creating Visio shape {0} of {1}: {2}", "step.vsdxShape"
    gStr.Add "Writing the Visio file ({0} native shapes, {1} pictures)", "step.vsdxWrite"
    gStr.Add "Writing the Visio file", "step.vsdxWrite0"
    gStr.Add "SVG via XPS: {0}", "step.xps"
    gStr.Add "Preparing a temporary slide (master and background hidden)", "step.xpsPrep"
    gStr.Add "PowerPoint is printing to XPS", "step.xpsPrint"
    gStr.Add "Packing the VSDX", "step.zip"
    gStr.Add ", transparency {0}%", "sum.alpha"
    gStr.Add "Background: {0}", "sum.bg"
    gStr.Add "PNG resolution: {0} DPI", "sum.dpi"
    gStr.Add "Default format: {0}", "sum.format"
    gStr.Add "Exports the selected shapes with the current settings", "sum.head"
    gStr.Add "Language: {0}", "sum.lang"
    gStr.Add "Margin: {0} pt", "sum.margin"
    gStr.Add "Several shapes: {0}", "sum.mode"
    gStr.Add "Saving: {0}", "sum.naming"
    gStr.Add "VSDX: native Visio shapes (editable text)", "sum.vsdxNative"
    gStr.Add "VSDX: embedded EMF vector picture", "sum.vsdxPicture"
End Sub

Private Sub LoadStrings_en_6()
    Dim s As String
    gStr.Add "conversion finished", "xps.done"
    gStr.Add "converting element {0} of {1}", "xps.elem"
    gStr.Add "embedding font {0}", "xps.font"
    gStr.Add "reading the XPS file", "xps.open"
    gStr.Add "analysing the page: {0} elements", "xps.page"
    gStr.Add "starting the converter", "xps.start"
    gStr.Add "writing the SVG file", "xps.write"
End Sub

Private Sub LoadStrings_ja_1()
    Dim s As String
    gStr.Add U("8868793A8A008A9EFF1A007B0030007DFF0830EA30DC30F3306E300C77E25370306B306430443066300D30B030EB30FC30D7306E8A008A9E30E130CB30E530FC30675207308A66FF30483089308C307E3059FF09"), "about.lang"
    s = ""
    s = s & U("77E253700020005600650063005300740061006D0070002000200076007B0030007D")
    s = s & vbCrLf
    s = s & U("90783093306762BC305930603051002020142014002056F35F62309200200045004D00460020002F00200053005600470020002F00200050004400460020002F002000560053004400580020002F00200050004E00470020306B30EF30F330AF30EA30C330AF306730A830AF30B930DD30FC30C8")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("4F5C8005FF1A007B0031007D")
    s = s & vbCrLf
    s = s & U("30E130FC30EBFF1A007B0032007D")
    s = s & vbCrLf
    s = s & U("30DB30FC30E030DA30FC30B8FF1A007B0033007D")
    s = s & vbCrLf
    s = s & U("30E930A430BB30F330B9FF1A004D004900540020004C006900630065006E00730065")
    s = s & vbCrLf
    s = s & "Copyright (c) {4} {1}"
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("77E25370306F305A3063306830AA30FC30D730F330BD30FC30B930FB71216599306730593002670965996A5F80FD30825E83544A3082306A304F300160C55831309253CE96C63059308B3053306830823042308A307E305B3093300230BD30FC30B930B330FC30C9306F305930793066002000470069007400480075006200203067516C958B305730663044307E30593002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("5F79306B7ACB3063305F3089300130EA30DC30F3306E300C77E25370306B306430443066300D30B030EB30FC30D7306B3042308B300C4F5C80053092652F63F4300D304B3089300100570065004300680061007400200050006100790020307E305F306F00200041006C0069007000610079002030674F5C8005306B30B330FC30D230FC309230543061305D3046305730663044305F30603051308B30683046308C30573044306730593002652F63F4306F4EFB610F306730016A5F80FD306B306F4E0052075F7197FF3057307E305B309330023044305F30603044305F652F63F4306F77E25370306E4FDD5B8830686539826F306B4F7F3044307E30593002")
    gStr.Add s, "about.text"
    gStr.Add U("007B0030007D0020306B306430443066"), "about.title"
    gStr.Add U("77E253700020005600650063005300740061006D0070"), "app.title"
    gStr.Add U("30D930FC30B830E5"), "bg.beige"
    gStr.Add U("30E630FC30B630FC8A2D5B9A306E8272"), "bg.custom"
    gStr.Add U("30E630FC30B630FC8A2D5B9A306E82720020007B0030007D"), "bg.customValue"
    gStr.Add U("65B9773C"), "bg.grid"
    gStr.Add U("65B9773CFF08959396940020007B0030007D002000700074FF09"), "bg.gridValue"
    gStr.Add U("900F660E"), "bg.none"
    gStr.Add U("767D"), "bg.white"
    gStr.Add U("30AD30E330F330BB30EB"), "btn.cancel"
    gStr.Add "OK", "btn.ok"
    gStr.Add U("8272306E5165529B4F8BFF1A003200350035002C003200340038002C0032003300350020307E305F306F00200023004600460046003800450042"), "col.bad"
    gStr.Add U("660E308B30550020002D007B0030007D0025"), "col.darker"
    gStr.Add U("8272309230AF30EA30C330AF305730669078629E300130C030D630EB30AF30EA30C330AF306778BA5B9A30020052002C0047002C00420020307E305F306F00200023005200520047004700420042002030825165529B3067304D307E30593002"), "col.hint"
    gStr.Add U("660E308B30550020002B007B0030007D0025"), "col.lighter"
    gStr.Add U("756A53F730925165529B305730668272309290783076304B30010052002C0047002C00420020002F00200023005200520047004700420042002030925165529B30573066304F306030553044FF1A"), "col.listPrompt"
    gStr.Add U("67008FD14F7F75283057305F8272"), "col.recent"
    gStr.Add U("67008FD14F7F75283057305F8272FF1A007B0030007D"), "col.recentItem"
    gStr.Add U("6A196E96306E8272"), "col.standard"
    gStr.Add U("6A196E96306E8272FF1A007B0030007D"), "col.stdItem"
    gStr.Add U("6FC330448D64007C8D64007C30AA30EC30F330B8007C9EC4007C858430447DD1007C7DD1007C858430449752007C9752007C6FC330449752007C7D2B"), "col.stdNames"
    gStr.Add U("30B730B930C630E0306E82722026"), "col.system"
    gStr.Add U("30B730B930C630E0306E8272"), "col.systemPicked"
    gStr.Add U("30C630FC30DE306E8272"), "col.theme"
    gStr.Add U("767D300180CC666F00200031007C9ED2300130C630AD30B930C800200031007C80CC666F00200032007C30C630AD30B930C800200032007C30A230AF30BB30F330C800200031007C30A230AF30BB30F330C800200032007C30A230AF30BB30F330C800200033007C30A230AF30BB30F330C800200034007C30A230AF30BB30F330C800200035007C30A230AF30BB30F330C800200036"), "col.themeNames"
    gStr.Add U("007B0030007D3001007B0031007D"), "col.tint"
    gStr.Add U("30E630FC30B630FC8A2D5B9A306E8272"), "col.title"
    gStr.Add U("8272306E5024"), "col.value"
    gStr.Add U("007B0030007DFF08007B0031007DFF09"), "col.withValue"
    gStr.Add U("30D530A130A430EB540DFF0862E15F355B50306A3057FF09FF1A"), "dlg.fileName"
    gStr.Add U("30A830AF30B930DD30FC30C8514830D530A930EB30C030FC30929078629E"), "dlg.pickFolder"
    gStr.Add U("30B930E930A430C930543068306E30D530A130A430EB306E4FDD5B58514830929078629EFF085F625F0FFF1A007B0030007DFF09"), "dlg.pickFolderBatch"
    gStr.Add U("30A830AF30B930DD30FC30C8514830D530A930EB30C030FC30929078629EFF0856F35F6230543068306B00200031002030D530A130A430EBFF09"), "dlg.pickFolderEach"
    gStr.Add U("007B0030007D002030683057306630A830AF30B930DD30FC30C8"), "dlg.saveAs"
    gStr.Add U("9078629E3057305F56F35F62309230A830AF30B930DD30FC30C8"), "dlg.saveTitle"
    gStr.Add U("0041006C0069007000610079FF08652F4ED85B9DFF09"), "don.alipay"
    s = ""
    s = s & U("77E253703092652F63F43057308830463068601D306330663044305F3060304D30013042308A304C30683046305430563044307E3059FF01")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("300C4F5C80053092652F63F4300D306E002000510052002030B330FC30C930920020007B0030007D002030678AAD307F53D630633066304F306030553044300291D1984D306F81EA7531306730593002")
    s = s & vbCrLf
    s = s & U("652F63F4306F4EFB610F3067300177E25370306E305930793066306E6A5F80FD306F305A306330687121659930FB30AA30FC30D730F330BD30FC30B9306730593002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("0047006900740048007500620020306730D730ED30B830A730AF30C8306B00200053007400610072002030924ED83051305F308A30015FC58981306A65B9306B7D394ECB3057305F308A305730663044305F30603051308B306E30825927304D306A52B1307F306B306A308A307E30593002")
    gStr.Add s, "don.thanks"
    gStr.Add U("652F63F4"), "don.title"
    gStr.Add U("005700650043006800610074FF085FAE4FE1FF09"), "don.wechat"
    gStr.Add U("30E630FC30B630FC8A2D5B9A"), "dpi.custom"
    gStr.Add U("9078629E3057305F56F35F62309230AF30EA30C330D730DC30FC30C97D4C7531306730B330D430FC3067304D307E305B309330673057305FFF084ED6306E30D730ED30B030E930E0304C4F7F75284E2D306E53EF80FD6027304C3042308A307E3059FF093002308230464E005EA6304A8A663057304F3060305530443002"), "err.clipboard"
    gStr.Add U("0045004D005A0020306E57277E2E306B593165573057307E3057305F3002"), "err.emz"
    gStr.Add U("9078629E3057305F30AA30D630B830A730AF30C8306F307E30683081306630A830AF30B930DD30FC30C83067304D307E305B30933002"), "err.group"
    gStr.Add U("006D00610063004F0053002088DC52A930B930AF30EA30D730C80020007B0030007D0020304C898B3064304B308A307E305B30933002006D00610063004F0053002030A430F330B930C830FC30E930FC3092308230464E005EA65B9F884C30573066304F3060305530443002"), "err.macHelper"
    gStr.Add U("FF0856F35F62309230A830AF30B930DD30FC30C83059308B306B306F00200050006F0077006500720050006F0069006E007400200066006F00720020004D00610063002000310036002E0038003200204EE5964D304C5FC5898130673059FF09"), "err.macVersion"
    gStr.Add U("30D530A130A430EB309266F8304D8FBC3081307E305B309330673057305FFF1A007B0030007D"), "err.macWrite"
    gStr.Add U("30D530A130A430EB304C4F5C62103055308C307E305B309330673057305FFF1A007B0030007D"), "err.noFile"
    gStr.Add U("0050004400460020306E30A830AF30B930DD30FC30C8306B593165573057307E3057305FFF1A007B0030007D"), "err.pdf"
End Sub

Private Sub LoadStrings_ja_2()
    Dim s As String
    gStr.Add U("00560053004400580020306E30D130C330B130FC30B85316306B593165573057307E3057305F3002"), "err.zip"
    gStr.Add U("00560053004400580020306E30D130C330B130FC30B85316306B593165573057307E3057305FFF0800570069006E0064006F0077007300204ED85C5E306E002000570069006E0064006F0077007300200050006F007700650072005300680065006C006C0020304C5FC5898130673059FF093002"), "err.zipPs"
    s = ""
    s = s & U("30D530A330FC30C930D030C330AF306E65B96CD530929078629E30573066304F306030553044FF1A")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("FF3B306F3044FF3D4F5C8005306B30E130FC30EB30929001308BFF08007B0030007DFF09")
    s = s & vbCrLf
    s = s & U("FF3B304430443048FF3D00470069007400480075006200203067002000490073007300750065002030924F5C6210FF08004700690074004800750062002030A230AB30A630F330C8304C5FC58981300251855BB9306F516C958B3055308C307E3059FF09")
    s = s & vbCrLf
    s = s & U("FF3B30AD30E330F330BB30EBFF3D623B308B")
    gStr.Add s, "fb.ask"
    s = ""
    s = s & U("002A002A554F984C0020002F002063D06848FF1A002A002A")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("002A002A518D73FE624B9806FF1A002A002A")
    s = s & vbCrLf
    s = s & "1. "
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("002A002A52D54F5C74B05883FF1A002A002A")
    s = s & vbCrLf
    s = s & "```"
    s = s & vbCrLf
    s = s & "{0}"
    s = s & vbCrLf
    s = s & "```"
    s = s & vbCrLf
    gStr.Add s, "fb.issueBody"
    gStr.Add U("005B30D530A330FC30C930D030C330AF005D0020"), "fb.issueTitle"
    s = ""
    s = s & U("4F5C8005306E30E130FC30EB30A230C930EC30B9309230AF30EA30C330D730DC30FC30C9306B30B330D430FC3057307E3057305FFF1A007B0030007D")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("4EFB610F306E30E130FC30EB30B530FC30D330B9FF08005700650062002030E130FC30EB3001004F00750074006C006F006F006B0020306A3069FF0930675B9B5148306B8CBC308A4ED83051306690014FE13067304D307E30593002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("65E25B9A306E30E130FC30EB30BD30D530C830924ECA30593050958B304D307E3059304BFF1F")
    s = s & vbCrLf
    s = s & U("FF0830E130FC30EB30BD30D530C830928A2D5B9A305730663044306A304458345408306F300C304430443048300D3092907830733001005700650062002030E130FC30EB304B308990014FE130573066304F306030553044FF09")
    gStr.Add s, "fb.mailCopied"
    s = ""
    s = s & U("4F5C8005306E30E130FC30EB30A230C930EC30B9FF1A007B0030007D")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("4EFB610F306E30E130FC30EB30B530FC30D330B9FF08005700650062002030E130FC30EB3001004F00750074006C006F006F006B0020306A3069FF09304B308990014FE13067304D307E30593002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("65E25B9A306E30E130FC30EB30BD30D530C830924ECA30593050958B304D307E3059304BFF1F")
    s = s & vbCrLf
    s = s & U("FF0830E130FC30EB30BD30D530C830928A2D5B9A305730663044306A304458345408306F300C304430443048300D3092907830733001005700650062002030E130FC30EB304B308990014FE130573066304F306030553044FF09")
    gStr.Add s, "fb.mailPlain"
    gStr.Add U("30E130FC30EB306730D530A330FC30C930D030C330AF"), "fb.mailTitle"
    s = ""
    s = s & U("30D630E930A630B630FC3092958B3051307E305B309330673057305F30026B21306E30DA30FC30B83092958B30443066304F306030553044FF1A")
    s = s & vbCrLf
    s = s & "{0}"
    gStr.Add s, "fb.noBrowser"
    gStr.Add U("30E130FC30EB30BD30D530C83092958B3051307E305B309330673057305F30026B21306E30A230C930EC30B9306B76F463A590014FE130573066304F306030553044FF1A007B0030007D"), "fb.noMailApp"
    gStr.Add U("77E253700020005600650063005300740061006D00700020007B0030007D002030D530A330FC30C930D030C330AF"), "fb.subject"
    gStr.Add U("30D530A330FC30C930D030C330AF"), "fb.title"
    gStr.Add U("30B930E930A430C9007B0030007D"), "file.page"
    gStr.Add "Finder", "fm.mac"
    gStr.Add U("30A830AF30B930D730ED30FC30E930FC"), "fm.win"
    gStr.Add U("0045004D0046002062E15F3530E130BF30D530A130A430EB"), "fmt.emf"
    gStr.Add U("0045004D005A002057277E2E62E15F3530E130BF30D530A130A430EB"), "fmt.emz"
    gStr.Add U("005000440046002030D930AF30BF30FC658766F8"), "fmt.pdf"
    gStr.Add U("0050004E004700209AD889E350CF5EA630D330C330C830DE30C330D7"), "fmt.png"
    gStr.Add U("005300560047002030B930B130FC30E930D630EB002030D930AF30BF30FC002030B030E930D530A330C330AF30B9"), "fmt.svg"
    gStr.Add U("005600530044005800200056006900730069006F002056F39762"), "fmt.vsdx"
    gStr.Add U("0057004D0046002000570069006E0064006F00770073002030E130BF30D530A130A430EB"), "fmt.wmf"
    s = ""
    s = s & U("77E253700020005600650063005300740061006D00700020306E4F7F304465B9")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("2460002056F35F62309200200031002030644EE54E0A9078629E30573066300C30A830AF30B930DD30FC30C8300D307E305F306F54045F625F0F306E30DC30BF30F3309230AF30EA30C330AF3057307E305930024F5530829078629E305730663044306A30443068304D306F30B930E930A430C951684F53309230A830AF30B930DD30FC30C83067304D307E30593002")
    s = s & vbCrLf
    s = s & U("2461002030B930E930A430C930543068306B30A830AF30B930DD30FC30C8FF1A00310020679A306B3064304D00200031002030D530A130A430EB300230B530E030CD30A430EB306730B930E930A430C93092890765709078629E3059308B306830019078629E3057305F30B930E930A430C930603051309230A830AF30B930DD30FC30C83057307E30593002")
    s = s & vbCrLf
    s = s & U("246200200045004D00460020753B50CF30B330D430FCFF1A9078629E3057305F56F35F62306E30D930AF30BF30FC753B50CF309296A3306B4F5C62103057307E3059FF085143306E56F35F62306F305D306E307E307EFF093002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("00B700200045004D00460020002F00200045004D005A0020002F00200057004D00460020002F0020005300560047FF1A30D930AF30BF30FC753B50CF30020050006F0077006500720050006F0069006E00740020003200300031003300204EE5964D306A30893069308C3067308200200053005600470020309230A830AF30B930DD30FC30C83067304D307E3059FF08003200330030003200203088308A524D306F51858535306E00200058005000530020590963DB30924F7F7528FF093002")
    s = s & vbCrLf
    s = s & U("00B70020005000440046FF1A56F35F623068540C30585927304D3055306E30D930AF30BF30FC00200050004400463002004C00610054006500580020306B67009069306730593002")
    s = s & vbCrLf
    s = s & U("00B700200056005300440058FF1A0056006900730069006F002030D530A130A430EB300256F35F62306F00200056006900730069006F0020306E30CD30A430C630A330D656F35F62306765875B5730927DE896C63067304D30010056006900730069006F0020306F4E0D8981306730593002")
    s = s & vbCrLf
    s = s & U("00B700200050004E0047FF1A0031003500300020002F00200033003000300020002F002000360030003000200044005000490020307E305F306F30E630FC30B630FC8A2D5B9AFF08670059270020007B0030007D0020004400500049FF093002")
    s = s & vbCrLf
    s = s & U("00B7002080CC666FFF1A300C80CC666F300D30E130CB30E530FC3067900F660E30FB767D30FB30D930FC30B830E530FB65B9773C30FB30E630FC30B630FC8A2D5B9A306E8272FF0830C630FC30DE306E82720020002F00206A196E96306E82720020002F0020305D306E4ED6306E82720020002F002030B930DD30A430C80020002F0020005200470042FF093092907830733001900F660E5EA630684F59767D30828A2D5B9A3067304D307E30593002")
    s = s & vbCrLf
    s = s & U("00B7002030A830AF30B930DD30FC30C84E2D306F9032884C72B66CC130A630A330F330C930A6306B73FE5728306E51E67406306852725408304C8868793A3055308C3001004500730063002030AD30FC30674E2D6B623067304D307E30593002")
    s = s & vbCrLf
    s = s & U("00B700208868793A8A008A9EFF1A300C77E25370306B306430443066300D30B030EB30FC30D7306E8A008A9E30E130CB30E530FC306700207B804F534E2D658730FB7E419AD44E2D658730FB0045006E0067006C00690073006830FB65E5672C8A9E002030925207308A66FF30483089308C307E3059FF0830593050306B53CD6620FF093002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("307E306830813066306E30A830AF30B930DD30FC30C8300180CC666F306E8FFD52A03001005000440046300130B930E930A430C930543068306E30A830AF30B930DD30FC30C83067306F300130AF30EA30C330D730DC30FC30C93068672B5C3E306E4E00664230B930E930A430C930924E0066427684306B4F7F75283057307E3059FF085B8C4E865F8C306B81EA52D53067524A9664FF093002")
    gStr.Add s, "help.ppt"
    gStr.Add U("003100203064306E30D530A130A430EB306B307E30683081308B"), "mode.merge"
    gStr.Add U("56F35F6230543068306B00200031002030D530A130A430EB"), "mode.split"
    gStr.Add U("305930793066306E30A830AF30B930DD30FC30C88A2D5B9A309265E25B9A5024306B623B3057307E3059304BFF1F"), "msg.askReset"
    s = ""
    s = s & U("3069306E30B930E930A430C9309230A830AF30B930DD30FC30C83057307E3059304BFF1F")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("306F3044FF1A9078629E3057305F0020007B0030007D0020679A30603051")
    s = s & vbCrLf
    s = s & U("304430443048FF1A51680020007B0031007D0020679A")
    gStr.Add s, "msg.askWhich"
    s = ""
    s = s & U("56F35F62304C9078629E3055308C30663044307E305B30933002")
    s = s & vbCrLf
    s = s & U("3053306E30B930E930A430C9306E305930793066306E56F35F62309230A830AF30B930DD30FC30C83057307E3059304BFF1F")
    gStr.Add s, "msg.askWholeSlide"
    gStr.Add U("80CC666F306E900F660E5EA6306B306F00200030FF5E0031003000300020306E65705024FF080025FF0930925165529B30573066304F3060305530443002"), "msg.badAlpha"
    gStr.Add U("89E350CF5EA6306B306F00200031FF5E007B0030007D0020306E65746570FF08004400500049FF0930925165529B30573066304F3060305530443002"), "msg.badDpi"
    gStr.Add U("65B9773C306E95939694306B306F00200032FF5E0032003000300020306E65705024FF0853584F4DFF1A00700074FF0930925165529B30573066304F3060305530443002"), "msg.badGrid"
    gStr.Add U("4F59767D306B306F00200030FF5E0035003000300020306E65705024FF0853584F4DFF1A00700074FF0930925165529B30573066304F3060305530443002"), "msg.badMargin"
    gStr.Add U("30B930E930A430C930543068306E30A830AF30B930DD30FC30C8306B593165573057307E3057305FFF1A007B0030007D"), "msg.batchFailed"
    gStr.Add U("007B0030007D002000440050004900203067306F7D040020007B0031007D002000D70020007B0032007D002030D430AF30BB30EBFF087D040020007B0033007D002030E130AC30D430AF30BB30EBFF09306B306A308A300166429593304C304B304B3063305F308A30E130E230EA304C4E0D8DB33057305F308A3059308B53EF80FD6027304C3042308A307E305930027D9A884C3057307E3059304BFF1F"), "msg.bigPng"
    gStr.Add U("30A830AF30B930DD30FC30C830924E2D6B623057307E3057305F3002"), "msg.cancelled"
    gStr.Add U("30A830AF30B930DD30FC30C830924E2D6B623057307E3057305F30025B8C4E863057305F0020007B0030007D0020500B306E30D530A130A430EB306F6B8B305730663042308A307E3059FF1A"), "msg.cancelledPartial"
    s = ""
    s = s & U("8272306E9078629E753B97623092958B3051307E305B309330673057305FFF1A007B0030007D")
    s = s & vbCrLf
    s = s & U("300C005200470042300D30DC30C330AF30B9306B8272309276F463A55165529B3059308B3053306830823067304D307E30593002")
    gStr.Add s, "msg.colorFailed"
    gStr.Add U("30A830AF30B930DD30FC30C8306B593165573057307E3057305FFF1A007B0030007D"), "msg.exportFailed"
    gStr.Add U("007B0030007D0020500B306E30D530A130A430EB309230A830AF30B930DD30FC30C83057307E3057305FFF1A"), "msg.exported"
    s = ""
    s = s & U("30B930DD30A430C830924F7F3046306B306F6A196E968868793A306730B930E930A430C93092958B304430663044308B5FC58981304C3042308A307E30593002")
    s = s & vbCrLf
    s = s & U("6A196E968868793A306B5207308A66FF30483066304B3089300C30B930DD30A430C8300D309230AF30EA30C330AF3057300130B930E930A430C94E0A306E8272309230AF30EA30C330AF30573066304F3060305530443002")
    gStr.Add s, "msg.eyedropperView"
    gStr.Add U("30D730ED30B830A730AF30C830DA30FC30B8FF1A007B0030007D"), "msg.home"
    gStr.Add U("307E306030A830AF30B930DD30FC30C8305730663044307E305B309330024E005EA630A830AF30B930DD30FC30C83059308B3068300130533053304B308967008FD14F7F75283057305F30D530A930EB30C030FC3092958B3051307E30593002"), "msg.noFolderYet"
    gStr.Add U("3053306E30B930E930A430C9306B306F30A830AF30B930DD30FC30C83067304D308B56F35F62304C3042308A307E305B30933002"), "msg.noShapesSlide"
    gStr.Add U("9078629E3057305F30B930E930A430C9306B306F30A830AF30B930DD30FC30C83067304D308B56F35F62304C3042308A307E305B30933002"), "msg.noShapesSlides"
    gStr.Add U("30D730EC30BC30F330C630FC30B730E730F3306B30B930E930A430C9304C3042308A307E305B30933002"), "msg.noSlides"
    s = ""
    s = s & U("005300560047002030924F5C62103067304D307E305B309330673057305F30023053306E00200050006F0077006500720050006F0069006E00740020306F00200053005600470020309276F463A530A830AF30B930DD30FC30C83067304D305A300151858535306E590963DB30826210529F3057307E305B309330673057305FFF0800570069006E0064006F0077007300204ED85C5E306E002000570069006E0064006F0077007300200050006F007700650072005300680065006C006C0020304C5FC5898130673059FF093002590963DB306B4F7F3048308B00200049006E006B00730063006100700065002030840020004C0069006200720065004F00660066006900630065002030823042308A307E305B30933002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("4EE3308F308A306B00200045004D00460020306730A830AF30B930DD30FC30C83057307E3059304BFF1FFF080050004400460020306730A830AF30B930DD30FC30C8305730664ED6306E30BD30D530C83067590963DB3059308B3053306830823067304D307E3059FF09")
    gStr.Add s, "msg.noSvg"
    gStr.Add U("6A196E968868793A306730A830AF30B930DD30FC30C83059308B56F35F6230929078629E30573066304F3060305530443002"), "msg.normalView"
    gStr.Add U("30D730EC30BC30F330C630FC30B730E730F33092958B304D300130A830AF30B930DD30FC30C83059308B56F35F6230929078629E30573066304F3060305530443002"), "msg.openAndSelect"
    gStr.Add U("5148306B30D730EC30BC30F330C630FC30B730E730F33092958B30443066304F3060305530443002"), "msg.openPres"
    gStr.Add U("0045004D00460020753B50CF30B330D430FC30924F5C62103067304D307E305B309330673057305FFF1A007B0030007D"), "msg.pictureFailed"
End Sub

Private Sub LoadStrings_ja_3()
    Dim s As String
    gStr.Add U("5148306B56F35F6230929078629E30573066304F3060305530443002"), "msg.selectFirst"
    gStr.Add U("7A7A767D306E30B930E930A430C90020007B0030007D0020679A309230B930AD30C330D73057307E3057305F3002"), "msg.skipped"
    s = ""
    s = s & U("51858535306E00200053005600470020590963DB304C0020007B0030007D002079D24EE55185306B7D42308F3089306A304B3063305F305F3081505C6B623057307E3057305F3002")
    s = s & vbCrLf
    s = s & vbCrLf
    s = s & U("56F35F62304C975E5E38306B890796D1306A58345408FF08592791CF306E65875B5730FB753B50CF30FB980270B9FF09306B8D773053308B30533068304C3042308A307E30593002")
    s = s & vbCrLf
    s = s & U("0049006E006B007300630061007000650020002F0020004C0069006200720065004F0066006600690063006500203067590963DBFF0830A430F330B930C830FC30EB6E08307F306E58345408FF093059308B304B30010045004D00460020306730A830AF30B930DD30FC30C83057307E3059304BFF1F")
    gStr.Add s, "msg.svgTimeout"
    gStr.Add U("6BCE56DE4FDD5B585148309290783076"), "naming.ask"
    gStr.Add U("65E25B9A306E30D530A930EB30C030FC306B81EA52D54FDD5B58"), "naming.auto"
    gStr.Add U("30B930E930A430C930543068306B30A830AF30B930DD30FC30C8FF08007B0030007D0020679A3001007B0031007DFF09"), "prog.batch"
    gStr.Add U("FF087D4C904E0020007B0030007D002079D2FF09"), "prog.elapsed"
    gStr.Add U("004500730063002030AD30FC30674E2D6B62"), "prog.escHint"
    gStr.Add U("007B0030007D0020309230A830AF30B930DD30FC30C8305730663044307E3059"), "prog.export1"
    gStr.Add U("007B0030007D0020500B306E0020007B0031007D002030D530A130A430EB309230A830AF30B930DD30FC30C8305730663044307E3059"), "prog.exportN"
    gStr.Add U("30B930DD30A430C8"), "prog.eyedropper"
    gStr.Add U("56F35F620020007B0030007D0020002F0020007B0031007DFF1A007B0032007D"), "prog.item"
    gStr.Add U("30B930E930A430C90020007B0030007D0020002F0020007B0031007DFF0830B930E930A430C9756A53F70020007B0032007DFF09"), "prog.page"
    gStr.Add U("6E965099305730663044307E30592026"), "prog.preparing"
    gStr.Add U("30D930FC30B830E5"), "rb.vsBgBeige"
    gStr.Add U("30E630FC30B630FC8A2D5B9A306E82722026"), "rb.vsBgCustomMac"
    gStr.Add U("65B9773CFF0830B030E930D575287D19FF09"), "rb.vsBgGrid"
    gStr.Add U("900F660E"), "rb.vsBgNone"
    gStr.Add U("30E630FC30B630FC8A2D5B9A306E8272"), "rb.vsBgSepCustom"
    gStr.Add U("767D"), "rb.vsBgWhite"
    gStr.Add U("77E25370306B306430443066"), "rb.vsBtnAbout"
    gStr.Add U("30D030FC30B830E730F330014F5C80053001004D00490054002030E930A430BB30F330B9300277E25370306F305A3063306830AA30FC30D730F330BD30FC30B930FB71216599306730593002"), "rb.vsBtnAbout.sup"
    gStr.Add U("77E253700020005600650063005300740061006D00700020306B306430443066"), "rb.vsBtnAbout.tip"
    gStr.Add U("30B930E930A430C930543068306B30A830AF30B930DD30FC30C8"), "rb.vsBtnBatch"
    gStr.Add U("540430B930E930A430C930920020003100203064306E30D530A130A430EB30683057306630A830AF30B930DD30FC30C83057307E3059FF0865E25B9A306E5F625F0FFF09300230B530E030CD30A430EB306730B930E930A430C93092890765709078629E3059308B30689078629E3057305F30B930E930A430C930603051309230A830AF30B930DD30FC30C83057307E305930027A7A767D306E30B930E930A430C930687A7A306E30BF30A430C830EB002030D730EC30FC30B930DB30EB30C030FC306F30B930AD30C330D73057307E30593002"), "rb.vsBtnBatch.sup"
    gStr.Add U("652F63F4306B3064304430662026"), "rb.vsBtnDonateInfo"
    gStr.Add U("62E15F3530E130BF30D530A130A430EB30020057006F007200640020002F00200057005000530020002F00200050006F0077006500720050006F0069006E0074002030674F7F304830018AD66587306E56F3306B67009069306730593002"), "rb.vsBtnEMF.sup"
    gStr.Add U("0045004D0046002030683057306630A830AF30B930DD30FC30C8"), "rb.vsBtnEMF.tip"
    gStr.Add U("0067007A006900700020306757277E2E3057305F00200045004D0046300251855BB9306F00200045004D004600203068540C30583067300130D530A130A430EB304C5C0F3055304F306A308A307E30593002"), "rb.vsBtnEMZ.sup"
    gStr.Add U("0045004D005A002030683057306630A830AF30B930DD30FC30C8"), "rb.vsBtnEMZ.tip"
    gStr.Add U("007B0030007D0020309230A830AF30B930DD30FC30C8"), "rb.vsBtnExport"
    gStr.Add U("9078629E3057305F56F35F62309230A830AF30B930DD30FC30C8"), "rb.vsBtnExport.tip"
    gStr.Add U("30B930DD30A430C8002800450029"), "rb.vsBtnEyedropper"
    gStr.Add U("30D530A330FC30C930D030C330AF"), "rb.vsBtnFeedback"
    gStr.Add U("4F5C8005306B30E130FC30EB30929001308B304B300100470069007400480075006200203067002000490073007300750065002030924F5C62103057307E30593002"), "rb.vsBtnFeedback.sup"
    gStr.Add U("554F984C306E5831544A30FB63D06848"), "rb.vsBtnFeedback.tip"
    gStr.Add U("65E25B9A306E30D530A930EB30C030FC2026"), "rb.vsBtnFolder"
    s = ""
    s = s & U("81EA52D54FDD5B583068300C30B930E930A430C930543068306B30A830AF30B930DD30FC30C8300D30674F7F304630D530A930EB30C030FCFF1A")
    s = s & vbCrLf
    s = s & "{0}"
    gStr.Add s, "rb.vsBtnFolder.sup"
    gStr.Add U("4F7F304465B9"), "rb.vsBtnHelp"
    gStr.Add U("30DC30BF30F330688A2D5B9A306E7C215358306A8AAC660E3002"), "rb.vsBtnHelp.sup"
    gStr.Add U("30D730ED30B830A730AF30C830DA30FC30B8"), "rb.vsBtnHome"
    gStr.Add U("30D630E930A630B630FC306777E25370306E0020004700690074004800750062002030DA30FC30B83092958B304D307E3059300265B03057304430D030FC30B830E730F3300130BD30FC30B930B330FC30C9300166F465B05C656B74309278BA8A8D3067304D307E30593002"), "rb.vsBtnHome.sup"
    gStr.Add U("73FE5728306E305930793066306E30A830AF30B930DD30FC30C88A2D5B9A30928868793A3057307E30593002"), "rb.vsBtnLauncher.sup"
    gStr.Add U("73FE5728306E8A2D5B9A30928868793A"), "rb.vsBtnLauncher.tip"
    gStr.Add U("305D306E4ED6306E82720028004D00292026"), "rb.vsBtnMoreColors"
    gStr.Add U("30A830AF30B930DD30FC30C851483092958B304F"), "rb.vsBtnOpenFolder"
    gStr.Add U("67005F8C306B30A830AF30B930DD30FC30C83057305F30D530A930EB30C030FC3092007B0030007D3067958B304D307E30593002"), "rb.vsBtnOpenFolder.sup"
    gStr.Add U("56F35F623068540C30585927304D3055306E30D930AF30BF30FC00200050004400463002"), "rb.vsBtnPDF.sup"
    gStr.Add U("005000440046002030683057306630A830AF30B930DD30FC30C8"), "rb.vsBtnPDF.tip"
    gStr.Add U("9AD889E350CF5EA630D330C330C830DE30C330D7300289E350CF5EA6306F300C30A830AF30B930DD30FC30C88A2D5B9A0020219200200050004E0047002089E350CF5EA6300D306790783073307E30593002"), "rb.vsBtnPNG.sup"
End Sub

Private Sub LoadStrings_ja_4()
    Dim s As String
    gStr.Add U("0050004E0047002030683057306630A830AF30B930DD30FC30C8"), "rb.vsBtnPNG.tip"
    gStr.Add U("0045004D00460020753B50CF30B330D430FC"), "rb.vsBtnPicture"
    gStr.Add U("9078629E3057305F56F35F62309230D930AF30BF30FC753B50CF30683057306696A3306B30B330D430FC3057300130EC30A430A230A630C83092300C56FA5B9A300D3057307E305930025143306E56F35F62306F305D306E307E307E306730593002"), "rb.vsBtnPicture.sup"
    gStr.Add U("8A2D5B9A309265E25B9A306B623B3059"), "rb.vsBtnReset"
    gStr.Add U("305930793066306E30A830AF30B930DD30FC30C88A2D5B9A306880CC666F306E8A2D5B9A309265E25B9A5024306B623B3057307E3059FF088868793A8A008A9E306F305D306E307E307EFF093002"), "rb.vsBtnReset.sup"
    gStr.Add U("30D630E930A630B630FC306776F463A5958B3051308B30B930B130FC30E930D630EB002030D930AF30BF30FC002030B030E930D530A330C330AF30B930020050006F0077006500720050006F0069006E00740020003200300031003300204EE5964D306A308930A830AF30B930DD30FC30C83067304D307E3059300265B03057304430D030FC30B830E730F3306F76F463A5300153E4304430D030FC30B830E730F3306F51858535306E00200058005000530020590963DB309281EA52D530674F7F75283057FF0865875B57306F65875B57306E307E307EFF09300130A830AF30B930DD30FC30C85F8C306B30D530A130A430EB309278BA8A8D3057307E30593002"), "rb.vsBtnSVG.sup"
    gStr.Add U("005300560047002030683057306630A830AF30B930DD30FC30C8"), "rb.vsBtnSVG.tip"
    gStr.Add U("0056006900730069006F002056F39762FF1A540456F35F62309200200056006900730069006F0020306E30CD30A430C630A330D656F35F62FF085F6272B630FB827230FB7DDA30FB77E2537030FB30B030EB30FC30D7FF09306B590963DB3057300165875B57306F305D306E307E307E7DE896C63067304D307E305930020056006900730069006F0020306F4E0D8981306730593002753B50CF308430B030E930D5306A3069590963DB3067304D306A304430AA30D630B830A730AF30C8306F30D930AF30BF30FC753B50CF30683057306657CB30818FBC307F307E30593002"), "rb.vsBtnVSDX.sup"
    gStr.Add U("0056005300440058002030683057306630A830AF30B930DD30FC30C8"), "rb.vsBtnVSDX.tip"
    gStr.Add U("5B8C4E865F8C306B30D530A930EB30C030FC3092958B304F"), "rb.vsChkOpen"
    gStr.Add U("30A830AF30B930DD30FC30C85F8C300130A830AF30B930D730ED30FC30E930FC0020002F002000460069006E0064006500720020306730D530A130A430EB30928868793A3057307E30593002"), "rb.vsChkOpen.sup"
    gStr.Add U("0056005300440058002030CD30A430C630A330D656F35F62"), "rb.vsChkVsdx"
    s = ""
    s = s & U("30AA30F3FF1A00560053004400580020306E56F35F62306F00200056006900730069006F0020306E30CD30A430C630A330D656F35F62306B306A308A300165875B573092305D306E307E307E7DE896C63067304D307E3059FF080056006900730069006F0020306F4E0D8981FF093002")
    s = s & vbCrLf
    s = s & U("30AA30D5FF1A56F35F6251684F5330920020003100203064306E00200045004D0046002030D930AF30BF30FC753B50CF30683057306657CB30818FBC307F30010050006F0077006500720050006F0069006E007400203068307E3063305F304F540C3058898B305F76EE306B306A308A307E30593002")
    gStr.Add s, "rb.vsChkVsdx.sup"
    gStr.Add U("0050004E0047002089E350CF5EA6"), "rb.vsDdDpi"
    gStr.Add U("0050004E00470020306E30A830AF30B930DD30FC30C830603051306B906975283055308C307E305930025B6688538A8C306E537052373067306F901A5E380020003300300030002000440050004900204EE54E0A30017DDA753B306F00200036003000300020004400500049300230E630FC30B630FC8A2D5B9AFF0867005927002000330030003000300020004400500049FF09308253EF80FD306730593002"), "rb.vsDdDpi.sup"
    gStr.Add U("65E25B9A306E5F625F0F"), "rb.vsDdFormat"
    gStr.Add U("300C30A830AF30B930DD30FC30C8300D3068300C30B930E930A430C930543068306B30A830AF30B930DD30FC30C8300D30674F7F30465F625F0FFF084FDD5B5830C030A430A230ED30B0306730824E0066427684306B590966F43067304D307E3059FF093002"), "rb.vsDdFormat.sup"
    gStr.Add U("89076570306E56F35F62"), "rb.vsDdMode"
    gStr.Add U("003100203064306E30D530A130A430EB306B307E30683081308B304B300156F35F6230543068306B30A830AF30B930DD30FC30C83057307E3059FF0830D530A130A430EB540D306F300C9078629E30A630A330F330C930A6300D306E56F35F62540DFF093002"), "rb.vsDdMode.sup"
    gStr.Add U("89076570306E56F35F6230929078629E3057305F3068304D"), "rb.vsDdMode.tip"
    gStr.Add U("4FDD5B5865B96CD5"), "rb.vsDdNaming"
    gStr.Add U("6BCE56DE30C030A430A230ED30B030674FDD5B585148309290783076304B300165E25B9A306E30D530A930EB30C030FC306B76F463A54FDD5B583057307E3059FF0830D530A130A430EB540D306B6642523B304C4ED8304D307E3059FF093002"), "rb.vsDdNaming.sup"
    gStr.Add U("900F660E5EA60025"), "rb.vsEbAlpha"
    gStr.Add U("00300020306F4E0D900F660E30010031003000300020306F5B8C5168306B900F660E300280CC666F30603051306B906975283055308C300156F35F62306B306F5F7197FF3057307E305B30933002"), "rb.vsEbAlpha.sup"
    gStr.Add U("80CC666F306E900F660E5EA6"), "rb.vsEbAlpha.tip"
    gStr.Add U("30E630FC30B630FC8A2D5B9A0020004400500049"), "rb.vsEbDpi"
    gStr.Add U("0031FF5E00330030003000300020306E657465703002300C0050004E0047002089E350CF5EA6300D304C300C30E630FC30B630FC8A2D5B9A300D306E3068304D306B4F7F3044307E3059300289E350CF5EA6304C9AD83044307B306930D530A130A430EB306F5927304D304F306A308A307E30593002"), "rb.vsEbDpi.sup"
    gStr.Add U("30E630FC30B630FC8A2D5B9A306E89E350CF5EA6"), "rb.vsEbDpi.tip"
    gStr.Add U("65B9773C0028007000740029"), "rb.vsEbGrid"
    gStr.Add U("65B9773C306E30DE30B976EE306E5927304D3055FF0853584F4DFF1A00700074FF0930020035002030DE30B930543068306B6FC330447DDA306B306A308A307E30593002"), "rb.vsEbGrid.sup"
    gStr.Add U("65B9773C306E95939694"), "rb.vsEbGrid.tip"
    gStr.Add U("4F59767D0028007000740029"), "rb.vsEbMargin"
    gStr.Add U("56F35F62306E5468308A306E4F59767DFF0853584F4DFF1A00700074300100310020007000740020224800200030002E003300350020006D006DFF09300200300020306F56F35F62306B30743063305F308A5408308F305B307E30593002"), "rb.vsEbMargin.sup"
    gStr.Add U("4F59767D"), "rb.vsEbMargin.tip"
    gStr.Add U("30E630FC30B630FC8A2D5B9A306E80CC666F8272"), "rb.vsEbRgb.tip"
    gStr.Add U("30E630FC30B630FC8A2D5B9A306E8272"), "rb.vsGalColor"
    s = ""
    s = s & U("73FE5728306E30E630FC30B630FC8A2D5B9A306E80CC666F8272FF1A007B0030007D")
    s = s & vbCrLf
    s = s & U("30C630FC30DE306E827230FB6A196E96306E827230FB67008FD14F7F75283057305F8272304B308990783076304B3001300C305D306E4ED6306E8272300D300C30B930DD30A430C8300D30924F7F3044307E30593002300C005200470042300D30DC30C330AF30B9306B00200052002C0047002C0042FF084F8BFF1A003200350035002C003200340038002C003200330035FF09307E305F306F002000310036002090326570FF084F8BFF1A0023004600460046003800450042FF0930925165529B3059308B3053306830823067304D307E305930028272309290783076306880CC666F306F300C30E630FC30B630FC8A2D5B9A306E8272300D306B5207308A66FF308F308A307E30593002")
    gStr.Add s, "rb.vsGalColor.sup"
    gStr.Add U("4F5C80053092652F63F4"), "rb.vsGalDonate"
    gStr.Add U("77E25370306F7121659930FB30AA30FC30D730F330BD30FC30B93067305930025F79306B7ACB3063305F308930010041006C00690070006100790020307E305F306F00200057006500430068006100740020005000610079002030674F5C8005306B30B330FC30D230FC309230543061305D3046305730663044305F30603051307E3059FF084EFB610F3067305930026A5F80FD306B306F4E0052075F7197FF3057307E305B3093FF093002"), "rb.vsGalDonate.sup"
    gStr.Add U("77E25370306B306430443066"), "rb.vsGrpAbout"
    gStr.Add U("80CC666F"), "rb.vsGrpBackground"
    gStr.Add U("9078629E7BC456F2309230A830AF30B930DD30FC30C8"), "rb.vsGrpExport"
    gStr.Add U("30A830AF30B930DD30FC30C88A2D5B9A"), "rb.vsGrpSettings"
    gStr.Add U("4E0062EC51E67406306830C430FC30EB"), "rb.vsGrpTools"
    gStr.Add U("80CC666FFF1A007B0030007D"), "rb.vsMnuBg"
    s = ""
    s = s & U("73FE5728306E80CC666FFF1A007B0030007D")
    s = s & vbCrLf
    s = s & U("30A830AF30B930DD30FC30C86642306B56F35F62306E4E0B306B6577304F80CC666FFF1A900F660E30FB767D30FB30D930FC30B830E530FB65B9773CFF0830B030E930D575287D19FF0930FB30E630FC30B630FC8A2D5B9A306E82723002300C30E630FC30B630FC8A2D5B9A306E8272300D3067306F30C630FC30DE306E827230FB6A196E96306E827230FB67008FD14F7F75283057305F8272304B308990783076304B3001300C305D306E4ED6306E8272300D300C30B930DD30A430C8300D30924F7F3048307E305930028272306E56DB89D2309230AF30EA30C330AF3059308B30688272306E9078629E753B9762304C958B304D307E30593002")
    gStr.Add s, "rb.vsMnuBg.sup"
    gStr.Add U("80CC666F"), "rb.vsMnuBg.tip"
    gStr.Add U("0057006F007200640020002F00200057005000530020002F00200050006F0077006500720050006F0069006E0074002030674F7F3048308B30D930AF30BF30FC753B50CF30028AD66587306E56F3306B67009069306730593002"), "rb.vsMnuEMF.desc"
    gStr.Add U("0067007A006900700020306757277E2E3057305F00200045004D0046300251855BB9306F540C3058306730B530A430BA304C5C0F3055304F306A308A307E30593002"), "rb.vsMnuEMZ.desc"
    gStr.Add U("73FE5728FF1A007B0030007D300277E25370306E8868793A8A008A9EFF087B804F534E2D658730FB7E419AD44E2D658730FB0045006E0067006C00690073006830FB65E5672C8A9EFF0930925207308A66FF3048307E3059300230593050306B53CD66203055308C307E30593002"), "rb.vsMnuLang.sup"
End Sub

Private Sub LoadStrings_ja_5()
    Dim s As String
    gStr.Add U("8868793A8A008A9E0020002F0020004C0061006E00670075006100670065"), "rb.vsMnuLang.tip"
    gStr.Add U("30DA30FC30B8304C56F35F623068540C30585927304D30553002004C00610054006500580020308462957A3F30B730B930C630E0306B67009069306730593002"), "rb.vsMnuPDF.desc"
    gStr.Add U("0031003500300020002F00200033003000300020002F002000360030003000200044005000490020307E305F306F30E630FC30B630FC8A2D5B9AFF0867005927002000330030003000300020004400500049FF093002"), "rb.vsMnuPNG.desc"
    gStr.Add U("00570065006230FB0049006C006C007500730074007200610074006F007230FB0049006E006B00730063006100700065002030674F7F3048307E305930020050006F0077006500720050006F0069006E00740020003200300031003300204EE5964D306A308930A830AF30B930DD30FC30C83067304D307E30593002"), "rb.vsMnuSVG.desc"
    gStr.Add U("658766F8"), "rb.vsMnuSepDoc"
    gStr.Add U("30D330C330C830DE30C330D7"), "rb.vsMnuSepRaster"
    gStr.Add U("30D930AF30BF30FC753B50CF"), "rb.vsMnuSepVector"
    gStr.Add U("0056006900730069006F0020306E30CD30A430C630A330D656F35F623067300165875B573092305D306E307E307E7DE896C63067304D307E305930020056006900730069006F0020306F4E0D8981306730593002"), "rb.vsMnuVSDX.desc"
    gStr.Add U("53E4304430BD30D530C8541130513002900F660E306830B030E930C730FC30B730E730F3306B306F5BFE5FDC3057307E305B30933002"), "rb.vsMnuWMF.desc"
    gStr.Add U("77E25370"), "rb.vsTab"
    gStr.Add U("73FE5728306E8A2D5B9A"), "set.current"
    gStr.Add U("FF08672A8A2D5B9AFF1A30D730EC30BC30F330C630FC30B730E730F3306E30D530A930EB30C030FC3002672A4FDD5B58307E305F306F30AF30E930A630C94E0A306E58345408306F30C730B930AF30C830C330D7FF09"), "set.folderNone"
    gStr.Add U("81EA52D54FDD5B58306B4F7F304665E25B9A306E30D530A930EB30C030FC30929078629E"), "set.pickDefaultFolder"
    gStr.Add U("56F35F62306E69CB9020309289E36790305730663044307E3059"), "step.analyze"
    gStr.Add U("4E00664230B930E930A430C9306756F35F6230FB80CC666F30FB4F59767D309254086210305730663044307E3059"), "step.compose"
    gStr.Add U("54086210304C5B8C4E863057307E3057305F"), "step.composed"
    gStr.Add U("5B8C4E86FF1A007B0030007D"), "step.doneFile"
    gStr.Add U("0049006E006B007300630061007000650020002F0020004C0069006200720065004F006600660069006300650020306700200045004D00460020309200200053005600470020306B590963DB305730663044307E3059"), "step.emf2svg"
    gStr.Add U("0045004D005A0020306B57277E2E305730663044307E3059FF080067007A00690070FF09"), "step.emz"
    gStr.Add U("30B930E930A430C94E0A306753D6308A8FBC30808272309230AF30EA30C330AF30573066304F306030553044FF0800360030002079D24EE55185FF09"), "step.eyedropper"
    gStr.Add U("0053005600470020309278BA8A8D30FB4FEE5FA9305730663044307E3059"), "step.fixSvg"
    gStr.Add U("0049006E006B007300630061007000650020304C00200045004D00460020309200200053005600470020306B590963DB305730663044307E3059"), "step.inkscape"
    gStr.Add U("004C0069006200720065004F006600660069006300650020304C00200045004D00460020309200200053005600470020306B590963DB305730663044307E3059"), "step.libreoffice"
    gStr.Add U("56F35F623068540C30585927304D3055306E4E006642658766F830924F5C6210305730663044307E3059"), "step.pdfDoc"
    gStr.Add U("30D930AF30BF30FC753B50CF3068305730668CBC308A4ED8305130663044307E3059"), "step.pdfPaste"
    gStr.Add U("0050006F0077006500720050006F0069006E00740020304C0020005000440046002030924F5C6210305730663044307E3059"), "step.pdfWrite"
    gStr.Add U("0050004E00470020309230EC30F330C030EA30F330B0305730663044307E3059FF1A007B0030007D002000D70020007B0031007D002030D430AF30BB30EBFF08007B0032007D0020004400500049FF09"), "step.png"
    gStr.Add U("0050006F0077006500720050006F0069006E00740020304C00200045004D0046002030D930AF30BF30FC753B50CF30924F5C6210305730663044307E3059"), "step.ppEmf"
    gStr.Add U("0050006F0077006500720050006F0069006E00740020304C00200053005600470020309230A830AF30B930DD30FC30C8305730663044307E3059"), "step.ppSvg"
    gStr.Add U("0050006F0077006500720050006F0069006E00740020304C00200057004D0046002030E130BF30D530A130A430EB30924F5C6210305730663044307E3059"), "step.ppWmf"
    gStr.Add U("56F35F6230926E965099305730663044307E3059FF08007B0030007D0020500B306E30AA30D630B830A730AF30C8FF09"), "step.prepare"
    gStr.Add U("0045004D0046002030D930AF30BF30FC753B50CF30924F5C6210305730663044307E3059FF08753B50CF30E230FC30C9FF09"), "step.vsdxPicture"
    gStr.Add U("0056006900730069006F002056F35F6230924F5C6210305730663044307E30590020007B0030007D0020002F0020007B0031007DFF1A007B0032007D"), "step.vsdxShape"
    gStr.Add U("0056006900730069006F002030D530A130A430EB309266F8304D51FA305730663044307E3059FF0830CD30A430C630A330D656F35F620020007B0030007D0020500B3001753B50CF0020007B0031007D0020679AFF09"), "step.vsdxWrite"
    gStr.Add U("0056006900730069006F002030D530A130A430EB309266F8304D51FA305730663044307E3059"), "step.vsdxWrite0"
    gStr.Add U("00580050005300207D4C753130670020005300560047002030924F5C6210FF1A007B0030007D"), "step.xps"
    gStr.Add U("4E00664230B930E930A430C930926E965099305730663044307E3059FF0830DE30B930BF30FC306880CC666F306F975E8868793AFF09"), "step.xpsPrep"
    gStr.Add U("0050006F0077006500720050006F0069006E00740020304C00200058005000530020306B53705237305730663044307E3059"), "step.xpsPrint"
    gStr.Add U("00560053004400580020309230D130C330B130FC30B85316305730663044307E3059"), "step.zip"
    gStr.Add U("3001900F660E5EA60020007B0030007D0025"), "sum.alpha"
    gStr.Add U("80CC666FFF1A007B0030007D"), "sum.bg"
    gStr.Add U("0050004E0047002089E350CF5EA6FF1A007B0030007D0020004400500049"), "sum.dpi"
    gStr.Add U("65E25B9A306E5F625F0FFF1A007B0030007D"), "sum.format"
    gStr.Add U("73FE5728306E8A2D5B9A30679078629E3057305F56F35F62309230A830AF30B930DD30FC30C83057307E3059"), "sum.head"
    gStr.Add U("8868793A8A008A9EFF1A007B0030007D"), "sum.lang"
    gStr.Add U("4F59767DFF1A007B0030007D002000700074"), "sum.margin"
    gStr.Add U("89076570306E56F35F62FF1A007B0030007D"), "sum.mode"
    gStr.Add U("4FDD5B5865B96CD5FF1A007B0030007D"), "sum.naming"
    gStr.Add U("0056005300440058FF1A0056006900730069006F0020306E30CD30A430C630A330D656F35F62FF0865875B5730927DE896C653EF80FDFF09"), "sum.vsdxNative"
    gStr.Add U("0056005300440058FF1A0045004D0046002030D930AF30BF30FC753B50CF309257CB30818FBC307F"), "sum.vsdxPicture"
End Sub

Private Sub LoadStrings_ja_6()
    Dim s As String
    gStr.Add U("590963DB304C5B8C4E863057307E3057305F"), "xps.done"
    gStr.Add U("89817D203092590963DB305730663044307E30590020007B0030007D0020002F0020007B0031007D"), "xps.elem"
    gStr.Add U("30D530A930F330C8309257CB30818FBC309330673044307E3059FF08007B0030007D0020500B76EEFF09"), "xps.font"
    gStr.Add U("005800500053002030D530A130A430EB30928AAD307F8FBC309330673044307E3059"), "xps.open"
    gStr.Add U("30DA30FC30B8309289E36790305730663044307E3059FF1A89817D200020007B0030007D0020500B"), "xps.page"
    gStr.Add U("590963DB30D730ED30B030E930E030928D7752D5305730663044307E3059"), "xps.start"
    gStr.Add U("005300560047002030D530A130A430EB309266F8304D51FA305730663044307E3059"), "xps.write"
End Sub

' <<< GENERATED
